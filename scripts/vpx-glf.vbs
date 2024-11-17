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
Dim glf_playerState : Set glf_playerState = CreateObject("Scripting.Dictionary")
Dim glf_running_shows : Set glf_running_shows = CreateObject("Scripting.Dictionary")
Dim glf_cached_shows : Set glf_cached_shows = CreateObject("Scripting.Dictionary")
Dim glf_lightPriority : Set glf_lightPriority = CreateObject("Scripting.Dictionary")
Dim glf_lightColorLookup : Set glf_lightColorLookup = CreateObject("Scripting.Dictionary")
Dim glf_lightMaps : Set glf_lightMaps = CreateObject("Scripting.Dictionary")
Dim glf_lightStacks : Set glf_lightStacks = CreateObject("Scripting.Dictionary")
Dim glf_lightTags : Set glf_lightTags = CreateObject("Scripting.Dictionary")
Dim glf_lightNames : Set glf_lightNames = CreateObject("Scripting.Dictionary")
Dim glf_modes : Set glf_modes = CreateObject("Scripting.Dictionary")
Dim glf_timers : Set glf_timers = CreateObject("Scripting.Dictionary")

Dim glf_ball_devices : Set glf_ball_devices = CreateObject("Scripting.Dictionary")
Dim glf_flippers : Set glf_flippers = CreateObject("Scripting.Dictionary")
Dim glf_ball_holds : Set glf_ball_holds = CreateObject("Scripting.Dictionary")
Dim glf_magnets : Set glf_magnets = CreateObject("Scripting.Dictionary")
Dim glf_segment_displays : Set glf_segment_displays = CreateObject("Scripting.Dictionary")


Dim bcpController : bcpController = Null
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = ""
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0

Dim glf_ballsPerGame : glf_ballsPerGame = 3
Dim glf_troughSize : glf_troughSize = tnob
Dim glf_lastTroughSw : glf_lastTroughSw = Null

Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False

Glf_RegisterLights()
Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController
    Set bcpController = (new GlfVpxBcpController)(bcpPort, bcpExeName)
End Sub

Public Sub Glf_Init()
	Glf_Options Null 'Force Options Check

	If glf_troughSize > 0 Then : swTrough1.DestroyBall : Set glf_ball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1) : Set glf_lastTroughSw = swTrough1 : End If
	If glf_troughSize > 1 Then : swTrough2.DestroyBall : Set glf_ball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2) : Set glf_lastTroughSw = swTrough2 : End If
	If glf_troughSize > 2 Then : swTrough3.DestroyBall : Set glf_ball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3) : Set glf_lastTroughSw = swTrough3 : End If
	If glf_troughSize > 3 Then : swTrough4.DestroyBall : Set glf_ball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4) : Set glf_lastTroughSw = swTrough4 : End If
	If glf_troughSize > 4 Then : swTrough5.DestroyBall : Set glf_ball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5) : Set glf_lastTroughSw = swTrough5 : End If
	If glf_troughSize > 5 Then : swTrough6.DestroyBall : Set glf_ball6 = swTrough6.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6) : Set glf_lastTroughSw = swTrough6 : End If
	If glf_troughSize > 6 Then : swTrough7.DestroyBall : Set glf_ball7 = swTrough7.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7) : Set glf_lastTroughSw = swTrough7 : End If
	If glf_troughSize > 7 Then : Drain.DestroyBall : Set glf_ball8 = Drain.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8) : End If
	
	Dim switch, switchHitSubs
	switchHitSubs = ""
	For Each switch in Glf_Switches
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_Hit() : DispatchPinEvent """ & switch.Name & "_active"", ActiveBall : End Sub" & vbCrLf
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_UnHit() : DispatchPinEvent """ & switch.Name & "_inactive"", ActiveBall : End Sub" & vbCrLf
	Next
	ExecuteGlobal switchHitSubs

	Dim slingshot, slingshotHitSubs
	slingshotHitSubs = ""
	For Each slingshot in Glf_Slingshots
		slingshotHitSubs = slingshotHitSubs & "Sub " & slingshot.Name & "_Slingshot() : DispatchPinEvent """ & slingshot.Name & "_active"", ActiveBall : End Sub" & vbCrLf
	Next
	ExecuteGlobal slingshotHitSubs

	Dim spinner, spinnerHitSubs
	spinnerHitSubs = ""
	For Each spinner in Glf_Spinners
		spinnerHitSubs = spinnerHitSubs & "Sub " & spinner.Name & "_Spin() : DispatchPinEvent """ & spinner.Name & "_active"", ActiveBall : End Sub" & vbCrLf
	Next
	ExecuteGlobal spinnerHitSubs

	If glf_debugEnabled = True Then

		' Calculate the scale factor
		Dim scaleFactor
		scaleFactor = 1080 / tableheight

		Dim light
		Dim monitorYaml : monitorYaml = "light:" & vbCrLf
		Dim godotLightScene : godotLightScene = ""
		For Each light in glf_lights
			monitorYaml = monitorYaml + "  " & light.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    size: 0.04" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& light.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& light.y/tableheight & vbCrLf


			godotLightScene = godotLightScene + "[node name="""&light.name&""" type=""Sprite2D"" parent=""lights""]" & vbCrLf
			godotLightScene = godotLightScene + "position = Vector2("&light.x*scaleFactor&", "&light.y*scaleFactor&")" & vbCrLf
			godotLightScene = godotLightScene + "script = ExtResource(""3_qb2nn"")" & vbCrLf
			godotLightScene = godotLightScene + "tags = []" & vbCrLf
			godotLightScene = godotLightScene + vbCrLf
		Next

		monitorYaml = monitorYaml + vbCrLf
		monitorYaml = monitorYaml + "switch:" & vbCrLf
		For Each switch in glf_switches
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& switch.y/tableheight & vbCrLf
		Next
		Dim troughCount
		For troughCount=1 to tnob
			monitorYaml = monitorYaml + "  s_trough" & troughCount & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& Eval("swTrough"&troughCount).x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& Eval("swTrough"&troughCount).y/tableheight & vbCrLf
		Next
		monitorYaml = monitorYaml + "  s_start:"&vbCrLf
		monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
		monitorYaml = monitorYaml + "    x: 0.95" & vbCrLf
		monitorYaml = monitorYaml + "    y: 0.95" & vbCrLf


		Dim fso, modesFolder, TxtFileStream, monitorFolder
		Set fso = CreateObject("Scripting.FileSystemObject")
		monitorFolder = "glf_mpf\monitor\"
		If Not fso.FolderExists("glf_mpf") Then
			fso.CreateFolder "glf_mpf"
		End If
		If Not fso.FolderExists("glf_mpf\monitor") Then
			fso.CreateFolder "glf_mpf\monitor"
		End If
		Set TxtFileStream = fso.OpenTextFile(monitorFolder & "\monitor.yaml", 2, True)
		TxtFileStream.WriteLine monitorYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(monitorFolder & "\gotdotlights.txt", 2, True)
		TxtFileStream.WriteLine godotLightScene
		TxtFileStream.Close
		
	End If

	'Cache Shows
	Dim mode, show_count, shot_count, cached_show
	show_count = 0
	shot_count = 0
	For Each mode in glf_modes.Items()
		glf_debugLog.WriteToLog "Init", mode.Name
		If Not IsNull(mode.ShowPlayer) Then
			With mode.ShowPlayer()
				Dim show_settings
				For Each show_settings in .EventShows()
					If Not IsNull(show_settings.Show) And show_settings.Action = "play" Then
						show_settings.InternalCacheId = CStr(show_count)
						show_count = show_count + 1
						glf_debugLog.WriteToLog "Show Settings", "show_player_" & mode.name & "_" & show_settings.Key & "_" & show_settings.InternalCacheId
						cached_show = Glf_ConvertShow(show_settings.Show, show_settings.Tokens)
						glf_cached_shows.Add "show_player_" & mode.name & "_" & show_settings.Key & "__" & show_settings.InternalCacheId, cached_show
					End If 
				Next
			End With
		End If

		If Not IsNull(mode.LightPlayer) Then
			With mode.LightPlayer()
				.ReloadLights()
			End With
		End If

		If UBound(mode.ModeShots) > -1 Then
			Dim mode_shot
			For Each mode_shot in mode.ModeShots
				Dim shot_profile : Set shot_profile = Glf_ShotProfiles(mode_shot.Profile)
				Dim x
				If mode_shot.InternalCacheId = -1 Then
					mode_shot.InternalCacheId = shot_count
					shot_count = shot_count + 1
				End If
				For x=0 to shot_profile.StatesCount
					Dim state
					Set state = shot_profile.StateForIndex(x)
					If state.InternalCacheId = -1 Then
						state.InternalCacheId = CStr(show_count)
						show_count = show_count + 1
					End If

					glf_debugLog.WriteToLog "Shot State", mode.name & "_" & x & "_" & mode_shot.Name & "_" & state.Key & "_" & mode_shot.InternalCacheId & "_" & state.InternalCacheId

					Dim key
					Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
					If Not IsNull(state.Tokens) Then
						For Each key In state.Tokens.Keys()
							mergedTokens.Add key, state.Tokens()(key)
						Next
					End If
					Dim tokens
					If Not IsNull(mode_shot.Tokens) Then
						Set tokens = mode_shot.Tokens
						For Each key In tokens.Keys
							If mergedTokens.Exists(key) Then
								mergedTokens(key) = tokens(key)
							Else
								mergedTokens.Add key, tokens(key)
							End If
						Next
					End If
					cached_show = Glf_ConvertShow(state.Show, mergedTokens)
					glf_cached_shows.Add mode.name & "_" & x & "_" & mode_shot.Name & "_" & state.Key & "_" & mode_shot.InternalCacheId & "_" & state.InternalCacheId, cached_show
				Next
			Next
		End If
	Next

	Glf_Reset()
End Sub

Sub Glf_Reset()

	DispatchQueuePinEvent "reset_complete", Null
End Sub

Sub Glf_Options(ByVal eventId)
	Dim ballsPerGame : ballsPerGame = Table1.Option("Balls Per Game", 1, 2, 1, 1, 0, Array("3 Balls", "5 Balls"))
	If ballsPerGame = 1 Then
		glf_ballsPerGame = 3
	Else
		glf_ballsPerGame = 5
	End If

	Dim glfDebug : glfDebug = Table1.Option("Glf Debug Log", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfDebug = 1 Then
		glf_debugEnabled = True
		glf_debugLog.EnableLogs
	Else
		glf_debugEnabled = False
		glf_debugLog.DisableLogs
	End If

	Dim glfuseBCP : glfuseBCP = Table1.Option("Glf Backbox Control Protocol", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseBCP = 1 Then
		useBCP = True
		If IsNull(bcpController) Then
			Glf_ConnectToBCPMediaController
		End If
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
	If glf_debugEnabled = True Then
		glf_debugLog.DisableLogs
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
		
		If keycode = LockbarKey Then
			DispatchPinEvent "s_lockbar_key_active", Null
		End If

		If KeyCode = PlungerKey Then
			DispatchPinEvent "s_plunger_key_active", Null
		End If

		If KeyCode = LeftMagnaSave Then
			DispatchPinEvent "s_left_magna_key_active", Null
		End If

		If KeyCode = RightMagnaSave Then
			DispatchPinEvent "s_right_magna_key_active", Null
		End If

		If KeyCode = StagedRightFlipperKey Then
			DispatchPinEvent "s_right_staged_flipper_key_active", Null
		End If

		If KeyCode = StagedLeftFlipperKey Then
			DispatchPinEvent "s_left_staged_flipper_key_active", Null
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
			DispatchPinEvent "s_plunger_key_inactive", Null
		End If

		If keycode = LeftFlipperKey Then
			DispatchPinEvent "s_left_flipper_inactive", Null
		End If
		
		If keycode = RightFlipperKey Then
			DispatchPinEvent "s_right_flipper_inactive", Null
		End If

		If keycode = LockbarKey Then
			DispatchPinEvent "s_lockbar_key_inactive", Null
		End If

		If KeyCode = LeftMagnaSave Then
			DispatchPinEvent "s_left_magna_key_inactive", Null
		End If

		If KeyCode = RightMagnaSave Then
			DispatchPinEvent "s_right_magna_key_inactive", Null
		End If

		If KeyCode = StagedRightFlipperKey Then
			DispatchPinEvent "s_right_staged_flipper_key_inactive", Null
		End If

		If KeyCode = StagedLeftFlipperKey Then
			DispatchPinEvent "s_left_staged_flipper_key_inactive", Null
		End If
		
	End If
End Sub

Dim glf_lastEventExecutionTime, glf_lastBcpExecutionTime, glf_lastLightUpdateExecutionTime
glf_lastEventExecutionTime = 0
glf_lastBcpExecutionTime = 0
glf_lastLightUpdateExecutionTime = 0

Public Sub Glf_GameTimer_Timer()

    'If (gametime - glf_lastEventExecutionTime) >= 33 Then
     '   glf_lastEventExecutionTime = gametime
		DelayTick
    'End If
	If (gametime - glf_lastBcpExecutionTime) >= 300 Then
        glf_lastBcpExecutionTime = gametime
		Glf_BcpUpdate
    End If

End Sub

Public Function Glf_RegisterLights()

	Dim light, tags, tag
	For Each light In Glf_Lights
		tags = Split(light.BlinkPattern, ",")
		For Each tag in tags
			
			tag = Trim(tag) ' Remove any leading or trailing spaces
			If Not glf_lightTags.Exists(tag) Then
				Set glf_lightTags(tag) = CreateObject("Scripting.Dictionary")
			End If
			If Not glf_lightTags(tag).Exists(light.Name) Then
				glf_lightTags(tag).Add light.Name, True
			End If
		Next
		glf_lightPriority.Add light.Name, 0
		
		Dim e, lmStr: lmStr = "lmArr = Array("    
		For Each e in GetElements()
			On Error Resume Next
			If InStr(LCase(e.Name), LCase("_" & light.Name & "_")) Then
				lmStr = lmStr & e.Name & ","
			End If
			For Each tag in tags
				If InStr(LCase(e.Name), LCase("_" & tag & "_")) Then
					lmStr = lmStr & e.Name & ","
				End If
			Next
			If Err Then Log "Error: " & Err
		Next
		lmStr = lmStr & "Null)"
		lmStr = Replace(lmStr, ",Null)", ")")
		ExecuteGlobal "Dim lmArr : "&lmStr
		glf_lightMaps.Add light.Name, lmArr
		glf_lightNames.Add light.Name, light
		Dim lightStack : Set lightStack = (new GlfLightStack)()
		glf_lightStacks.Add light.Name, lightStack
		light.State = 1
		Glf_SetLight light.Name, "000000"
	Next
End Function

Public Function Glf_SetLight(light, color)
	
	Dim rgbColor
	If glf_lightColorLookup.Exists(color) Then
		rgbColor = glf_lightColorLookup(color)
	Else
		glf_lightColorLookup.Add color, RGB( CInt("&H" & (Left(color, 2))), CInt("&H" & (Mid(color, 3, 2))), CInt("&H" & (Right(color, 2)) ))
		rgbColor = glf_lightColorLookup(color)
	End If
	

	If IsNull(color) Then
		glf_lightNames(light).Color = rgb(0,0,0)
	Else
		glf_lightNames(light).Color = rgbColor
	End If
	dim lightMap
	For Each lightMap in glf_lightMaps(light)
		If Not IsNull(lightMap) Then
			On Error Resume Next
			lightMap.Color = glf_lightNames(light).Color
			If Err Then Debug.Print "Error: " & Err & ". Light:" & light & ", LightMap: " & lightMap.Name
		End If
	Next
End Function

Public Function Glf_ParseInput(value)
	Dim templateCode : templateCode = ""
	Dim tmp: tmp = value
	Dim isVariable, parts
    Select Case VarType(value)
        Case 8 ' vbString
			tmp = Glf_ReplaceCurrentPlayerAttributes(tmp)
			tmp = Glf_ReplaceDeviceAttributes(tmp)
			'msgbox tmp
			If InStr(tmp, " if ") Then
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & Glf_ConvertIf(tmp, "Glf_" & glf_FuncCount) & vbCrLf
				templateCode = templateCode & "End Function"
			Else
				isVariable = Glf_IsCondition(tmp)
				If Not IsNull(isVariable) Then
					'The input needs formatting
					parts = Split(isVariable, ":")
					If UBound(parts) = 1 Then
						tmp = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
					End If
				End If
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
				templateCode = templateCode & "End Function"
			End IF
        Case Else
			templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf			
			isVariable = Glf_IsCondition(tmp)
			If Not IsNull(isVariable) Then
				'The input needs formatting
				parts = Split(isVariable, ":")
				If UBound(parts) = 1 Then
					tmp = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
				End If
			End If
			templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
			templateCode = templateCode & "End Function"
    End Select
	'msgbox templateCode
	ExecuteGlobal templateCode
	Dim funcRef : funcRef = "Glf_" & glf_FuncCount
	glf_FuncCount = glf_FuncCount + 1
	Glf_ParseInput = Array(funcRef, value, True)
End Function

Public Function Glf_ParseEventInput(value)
	Dim templateCode : templateCode = ""
	Dim condition : condition = Glf_IsCondition(value)
	If IsNull(condition) Then
		Glf_ParseEventInput = Array(value, value, Null)
	Else
		dim conditionReplaced : conditionReplaced = Glf_ReplaceCurrentPlayerAttributes(condition)
		conditionReplaced = Glf_ReplaceDeviceAttributes(conditionReplaced)
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

Function Glf_ReplaceDeviceAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "devices\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
	replacement = "glf_$1(""$2"").GetValue(""$3"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceDeviceAttributes = outputString
End Function

Function Glf_ConvertIf(value, retName)
    Dim parts, condition, truePart, falsePart, isVariable
    parts = Split(value, " if ")
    truePart = Trim(parts(0))
    Dim conditionAndFalsePart
    conditionAndFalsePart = Split(parts(1), " else ")
    condition = Trim(conditionAndFalsePart(0))
    falsePart = Trim(conditionAndFalsePart(1))
	isVariable = Glf_IsCondition(truePart)
	If Not IsNull(isVariable) Then
		'The input needs formatting
		parts = Split(isVariable, ":")
		If UBound(parts) = 1 Then
			truePart = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
		End If
	End If
	isVariable = Glf_IsCondition(falsePartPart)
	If Not IsNull(isVariable) Then
		'The input needs formatting
		parts = Split(isVariable, ":")
		If UBound(parts) = 1 Then
			falsePart = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
		End If
	End If

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
	value = Replace(value, "&&", "And")
	Glf_ConvertCondition = "    "&retName&" = " & value
End Function

Function Glf_FormatValue(value, formatString)
    Dim padChar, width, result, align

    ' Default values
    padChar = " " ' Default padding character is space
    align = ">"   ' Default alignment is right
    width = 0     ' Default width is 0 (no padding)

    If Len(formatString) >= 2 Then
        padChar = Mid(formatString, 1, 1)
        align = Mid(formatString, 2, 1)
        width = CInt(Mid(formatString, 3))
    End If

    Select Case align
        Case ">" ' Right-align with padding
            If Len(value) < width Then
                result = String(width - Len(value), padChar) & value
            Else
                result = value
            End If
        Case "<" ' Left-align with padding
            If Len(value) < width Then
                result = value & String(width - Len(value), padChar)
            Else
                result = value
            End If
        Case "^" ' Center-align with padding
            Dim leftPad, rightPad
            If Len(value) < width Then
                leftPad = (width - Len(value)) \ 2
                rightPad = width - Len(value) - leftPad
                result = String(leftPad, padChar) & value & String(rightPad, padChar)
            Else
                result = value
            End If
        Case Else ' Default: Return value as is
            result = value
    End Select

    Glf_FormatValue = result
End Function


Sub glf_ConvertYamlShowToGlfShow(yamlFilePath)
    Dim fso, file, content, lines, line, output, i, stepLights
    Dim glf_ShowName, stepTime, lightsDict, key, lightName, color, intensity
    
    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Read the YAML file
    Set file = fso.OpenTextFile(yamlFilePath, 1)
    content = file.ReadAll
    file.Close
    
    ' Split the content into lines
    lines = Split(content, vbLf)
    
    ' Initialize variables
    glf_ShowName = fso.GetBaseName(yamlFilePath)
    output = "Dim glf_Show" & glf_ShowName & " : Set glf_Show" & glf_ShowName & " = (new GlfShow)(""" & glf_ShowName & """)" & vbCrLf
    output = output & "With glf_Show" & glf_ShowName & vbCrLf
    
    ' Iterate through lines to extract steps and lights
	stepLights = ""
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Close the step when a new time step or end of file is reached
        If InStr(line, "- time:") = 1 And stepLights <> "" Then
			'output = output & vbTab & vbTab & ".Lights = Array("& Left(stepLights, Len(stepLights) - 1)&")" & vbCrLf
			output = output & vbTab & vbTab & ".Lights = Array(" & _ 
            SplitStringWithUnderscore(Left(stepLights, Len(stepLights) - 1), 1500) & ")" & vbCrLf
            output = output & vbTab & "End With" & vbCrLf
			stepLights = ""
		End If

        ' Identify time steps
        If InStr(line, "- time:") = 1 Then
            stepTime = Trim(Split(line, ":")(1))
            output = output & vbTab & "With .AddStep(" & stepTime & ", Null, Null)" & vbCrLf
        
		ElseIf InStr(line, "lights:") = 1 Then
			'Do Nothing
        ' Identify lights and colors
        ElseIf InStr(line, "l") = 1 Then
            key = Split(line, ":")(0)
            lightName = Trim(key)
            color = Trim(Split(line, """")(1))
            
            ' Default intensity to 100 if color is not "000000"
            intensity = 100
            'If color = "000000" Then
            '    intensity = 0
            'End If
            
            ' Add lights to output
            'If intensity > 0 Then
                
            'End If
			stepLights = stepLights + """" & lightName & "|" & intensity & "|" & color & ""","

        End If
    Next
    
    ' Close the final step and the show
	'output = output & vbTab & vbTab & ".Lights = Array("& Left(stepLights, Len(stepLights) - 1)&")" & vbCrLf
	output = output & vbTab & vbTab & ".Lights = Array(" & _ 
    SplitStringWithUnderscore(Left(stepLights, Len(stepLights) - 1), 1500) & ")" & vbCrLf
    output = output & vbTab & "End With" & vbCrLf
    output = output & "End With" & vbCrLf
    
    ' Write the output to a VBScript file
    Dim outputFilePath
    outputFilePath = fso.GetParentFolderName(yamlFilePath) & "\" & glf_ShowName & ".vbs"
    Set file = fso.CreateTextFile(outputFilePath, True)
    file.Write output
    file.Close
    
    ' Clean up
    Set fso = Nothing
    Set file = Nothing
End Sub

Function SplitStringWithUnderscore(str, maxLength)
    Dim result, i, strLen
    strLen = Len(str)
    result = ""
    
    If strLen <= maxLength Then
        result = str
    Else
        For i = 1 To strLen Step maxLength
            If i + maxLength - 1 < strLen Then
                result = result & Mid(str, i, maxLength) & "_" & vbCrLf & vbTab & vbTab & vbTab
            Else
                result = result & Mid(str, i, maxLength)
            End If
        Next
    End If
    
    SplitStringWithUnderscore = result
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
	If Not glf_modes.Exists(name) Then 
		Dim mode : Set mode = (new Mode)(name, priority)
		glf_modes.Add name, mode
		Set CreateGlfMode = mode
	End If
End Function

Function GlfModes(name)
	If glf_modes.Exists(name) Then 
		Set GlfModes = glf_modes(name)
	Else
		GlfModes = Null
	End If
End Function

Function GlfKwargs()
	Set GlfKwargs = CreateObject("Scripting.Dictionary")
End Function

Function Glf_ConvertShow(show, tokens)

	Dim showStep, light, lightsCount, x,tagLight, tagLights, lightParts, token, stepIdx
	Dim newShow, lightsInShow
	Set lightsInShow = CreateObject("Scripting.Dictionary")

	ReDim newShow(UBound(show.Steps().Keys()))
	stepIdx = 0
	For Each showStep in show.Steps().Items()
		lightsCount = 0 
		For Each light in showStep.Lights
			lightParts = Split(light, "|")
			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And Not glf_lightNames.Exists(lightParts(0)) Then
					tagLights = glf_lightTags(lightParts(0)).Keys()
					lightsCount = UBound(tagLights)+1
				Else
					If IsNull(token) Then
						lightsCount = lightsCount + 1
					Else
						'resolve token lights
						If Not glf_lightNames.Exists(tokens(token)) Then
							'token is a tag
							tagLights = glf_lightTags(tokens(token)).Keys()
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
		For Each light in showStep.Lights
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
				If IsNull(token) And Not glf_lightNames.Exists(lightParts(0)) Then
					tagLights = glf_lightTags(lightParts(0)).Keys()
					For Each tagLight in tagLights
						If UBound(lightParts) >=1 Then
							seqArray(x) = tagLight & "|"&lightParts(1)&"|"&lightColor
						Else
							seqArray(x) = tagLight & "|"&lightParts(1)
						End If
						If Not lightsInShow.Exists(tagLight) Then
							lightsInShow.Add tagLight, True
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
						If Not lightsInShow.Exists(lightParts(0)) Then
							lightsInShow.Add lightParts(0), True
						End If
						x=x+1
					Else
						'resolve token lights
						If Not glf_lightNames.Exists(tokens(token)) Then
							'token is a tag
							tagLights = glf_lightTags(tokens(token)).Keys()
							For Each tagLight in tagLights
								If UBound(lightParts) >=1 Then
									seqArray(x) = tagLight & "|"&lightParts(1)&"|"&lightColor
								Else
									seqArray(x) = tagLight & "|"&lightParts(1)
								End If
								If Not lightsInShow.Exists(tagLight) Then
									lightsInShow.Add tagLight, True
								End If
								x=x+1
							Next
						Else
							If UBound(lightParts) >= 1 Then
								seqArray(x) = tokens(token) & "|"&lightParts(1)&"|"&lightColor
							Else
								seqArray(x) = tokens(token) & "|"&lightParts(1)
							End If
							If Not lightsInShow.Exists(tokens(token)) Then
								lightsInShow.Add tokens(token), True
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
	Glf_ConvertShow = Array(newShow, lightsInShow)
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

Function CreateGlfInput(value)
	Set CreateGlfInput = (new GlfInput)(value)
End Function

Class GlfInput
	Private m_raw, m_value, m_isGetRef
  
    Public Property Get Value() 
		If m_isGetRef = True Then
			Value = GetRef(m_value)()
		Else
			Value = m_value
		End If
	End Property

    Public Property Get Raw() : Raw = m_raw : End Property

	Public default Function init(input)
        m_raw = input
        Dim parsedInput : parsedInput = Glf_ParseInput(input)
        m_value = parsedInput(0)
        m_isGetRef = parsedInput(2)
	    Set Init = Me
	End Function

End Class

'******************************************************
'*****   GLF Shows 		                           ****
'******************************************************

Dim glf_ShowOn : Set glf_ShowOn = (new GlfShow)("on")
With glf_ShowOn
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100")
	End With
End With

Dim glf_ShowOff : Set glf_ShowOff = (new GlfShow)("off")
With glf_ShowOff
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

Dim glf_ShowFlash : Set glf_ShowFlash = (new GlfShow)("flash")
With glf_ShowFlash
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

Dim glf_ShowFlashColor : Set glf_ShowFlashColor = (new GlfShow)("flash_color")
With glf_ShowFlashColor
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|(color)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

Dim glf_ShowOnColor : Set glf_ShowOnColor = (new GlfShow)("led_color")
With glf_ShowOnColor
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|(color)")
	End With
End With

Dim glf_Showtest : Set glf_Showtest = (new GlfShow)("test")
With glf_Showtest
	With .AddStep(0.00000, Null, Null)
		.Lights = Array("l142|100|000000","l143|100|000000","l141|100|000000","l140|100|00c737","l79|100|000000","l78|100|000000","l77|100|000000","l76|100|00d628","l75|100|00d628","l74|100|00d628","l73|100|00d32b","l72|100|000000","l71|100|000000","l70|100|000000","l69|100|000000","l44|100|000000","l11|100|000000","l12|100|000000","l13|100|000000","l14|100|000000","l15|100|000000","l05|100|000000","l07|100|000000","l08|100|000000","l09|100|000000","l10|100|000000","l06|100|000000","l04|100|000000","l03|100|000000","l02|100|000000","l01|100|000000","l16|100|000000","l23|100|000000","l63|100|000000","l53|100|000000","l68|100|00eb13","l67|100|00df1f","l66|100|00d727","l35|100|000000","l62|100|000000","l34|100|000000","l92|100|000000","l60|100|000000","l61|100|000000","l19|100|000000","l18|100|000000","l17|100|000000","l49|100|000000","l59|100|000000","l65|100|000000","l33|100|000000","l32|100|000000","l31|100|000000","l30|100|000000","l29|100|000000","l41|100|000000","l40|100|000000","l39|100|000000","l38|100|000000","l37|100|000000","l28|100|000000","l27|100|000000","l26|100|000000","l25|100|000000","l24|100|000000","l45|100|000000","l43|100|000000","l50|100|000000","l51|100|000000","l52|100|000000","l21|100|000000","l22|100|000000","l20|100|000000","l58|100|000000","l57|100|000000","l56|100|000000","l55|100|000000","l54|100|000000","l48|100|000000","l46|100|000000","l47|100|000000","l64|100|000000","l42|100|000000","l84|100|000000","l83|100|000000","l82|100|000000","l81|100|000000","l80|100|000000",_
			"l95|100|000000","l139|100|000000","l137|100|000000","l135|100|000000","l136|100|000000","l145|100|000000","l134|100|000000","l133|100|000000","l131|100|000000","l130|100|000000","l129|100|000000","l128|100|000000","l127|100|000000","l126|100|000000","l125|100|000000","l124|100|000000","l123|100|000000","l122|100|00ea14","l121|100|000000","l120|100|000000","l119|100|000000","l116|100|000000","l115|100|00ad50","l114|100|000000","l113|100|00d12d","l112|100|000000","l111|100|000000","l110|100|000000","l109|100|000000","l108|100|000000","l107|100|000000","l106|100|000000","l105|100|000000","l104|100|000000","l103|100|000000","l102|100|000000","l101|100|000000","l100|100|000000","l118|100|000000","l117|100|000000","l132|100|000000","l94|100|000000","l97|100|000000","l93|100|000000","l98|100|000000")
	End With
	With .AddStep(0.03333, Null, Null)
		.Lights = Array("l140|100|00c33b","l76|100|00d22c","l75|100|00d22c","l74|100|00d22c","l73|100|00cf2f","l68|100|00e717","l67|100|00db23","l66|100|00d32b","l130|100|00fd02","l122|100|00e618","l115|100|00a955","l113|100|00cd31","l97|100|00fd02")
	End With
	With .AddStep(0.06667, Null, Null)
		.Lights = Array("l140|100|00b945","l76|100|00c737","l75|100|00c737","l74|100|00c737","l73|100|00c539","l68|100|00dc22","l67|100|00d12d","l66|100|00c836","l131|100|00fa04","l130|100|00f20c","l122|100|00db23","l115|100|009e60","l113|100|00c23c","l97|100|00f20c")
	End With
	With .AddStep(0.10000, Null, Null)
		.Lights = Array("l140|100|00a955","l76|100|00b846","l75|100|00b846","l74|100|00b846","l73|100|00b549","l68|100|00cd31","l67|100|00c13d","l66|100|00b945","l131|100|00ea14","l130|100|00e21c","l129|100|00ff00","l127|100|00f509","l123|100|00f10c","l122|100|00cc32","l115|100|008e70","l113|100|00b34b","l97|100|00e21c")
	End With
	With .AddStep(0.13333, Null, Null)
		.Lights = Array("l140|100|009569","l76|100|00a35b","l75|100|00a35b","l74|100|00a35b","l73|100|00a15d","l68|100|00b944","l67|100|00ae50","l66|100|00a45a","l131|100|00d727","l130|100|00cf2f","l129|100|00ec12","l127|100|00e21c","l123|100|00de20","l122|100|00b846","l115|100|007b83","l113|100|009f5f","l97|100|00cf2f")
	End With
	With .AddStep(0.16667, Null, Null)
		.Lights = Array("l140|100|007f7f","l76|100|008d71","l75|100|008d71","l74|100|008d71","l73|100|008a74","l68|100|00a25c","l67|100|009668","l66|100|008e70","l139|100|00f10d","l131|100|00c03d","l130|100|00b846","l129|100|00d628","l127|100|00cb33","l123|100|00c836","l122|100|00a15d","l116|100|00ef0f","l115|100|006599","l114|100|00ea14","l113|100|008876","l97|100|00b846")
	End With
	With .AddStep(0.20000, Null, Null)
		.Lights = Array("l141|100|00fc03","l140|100|006698","l76|100|00748a","l75|100|00748a","l74|100|00748a","l73|100|00718d","l68|100|008975","l67|100|007d81","l66|100|007589","l139|100|00d826","l135|100|00ea14","l131|100|00a757","l130|100|009e60","l129|100|00bd41","l127|100|00b24c","l123|100|00af4f","l122|100|008876","l116|100|00d628","l115|100|004bb3","l114|100|00d12d","l113|100|006f8f","l97|100|009e60")
	End With
	With .AddStep(0.23333, Null, Null)
		.Lights = Array("l141|100|00e01e","l140|100|0049b4","l77|100|00f608","l76|100|0059a5","l75|100|0059a5","l74|100|0059a5","l73|100|0056a8","l68|100|006e90","l67|100|00629c","l66|100|005aa4","l139|100|00bd41","l135|100|00cf2f","l131|100|008b72","l130|100|00837b","l129|100|00a15d","l127|100|009668","l123|100|00936b","l122|100|006d91","l121|100|00fc02","l116|100|00bb43","l115|100|0030ce","l114|100|00b648","l113|100|0054aa","l97|100|00837b")
	End With
	With .AddStep(0.26667, Null, Null)
		.Lights = Array("l141|100|00c43a","l140|100|002dd1","l79|100|00e31a","l77|100|00d925","l76|100|003bc3","l75|100|003bc3","l74|100|003bc3","l73|100|0039c5","l68|100|0050ae","l67|100|0044ba","l66|100|003cc2","l139|100|009f5f","l135|100|00b24c","l131|100|006f8f","l130|100|006698","l129|100|00847a","l127|100|007985","l123|100|007688","l122|100|004faf","l121|100|00df1f","l116|100|009d61","l115|100|0013eb","l114|100|009866","l113|100|0036c8","l132|100|00ff00","l97|100|006698")
	End With
	With .AddStep(0.30000, Null, Null)
		.Lights = Array("l141|100|00a45a","l140|100|000fef","l79|100|00c539","l77|100|00bb43","l76|100|001de1","l75|100|001de1","l74|100|001de1","l73|100|001ae4","l23|100|00f30b","l68|100|0032cc","l67|100|0026d8","l66|100|001ee0","l139|100|00817d","l135|100|00936b","l131|100|0050ae","l130|100|0047b7","l129|100|006698","l127|100|005ba3","l123|100|0058a6","l122|100|0031cd","l121|100|00c13d","l116|100|007f7f","l115|100|000000","l114|100|007a84","l113|100|0018e6","l112|100|00ee10","l132|100|00e01d","l97|100|0047b7")
	End With
	With .AddStep(0.33333, Null, Null)
		.Lights = Array("l141|100|008579","l140|100|000000","l79|100|00a559","l77|100|009a63","l76|100|000000","l75|100|000000","l74|100|000000","l73|100|000000","l23|100|00d42a","l68|100|0013eb","l67|100|0007f7","l66|100|0000ff","l61|100|00e717","l139|100|00629c","l135|100|00738a","l131|100|0030ce","l130|100|0028d6","l129|100|0046b8","l127|100|003bc3","l124|100|00f30b","l123|100|0037c6","l122|100|0012ec","l121|100|00a05e","l116|100|005f9f","l114|100|005ba3","l113|100|000000","l112|100|00ce30","l111|100|00ee10","l132|100|00c13d","l94|100|00e11d","l97|100|0028d6")
	End With
	With .AddStep(0.36667, Null, Null)
		.Lights = Array("l141|100|006599","l79|100|008579","l78|100|00e816","l77|100|007a84","l23|100|00b44a","l68|100|000000","l67|100|000000","l66|100|000000","l61|100|00c737","l17|100|00f10c","l139|100|0041bd","l135|100|0053ab","l131|100|0010ee","l130|100|0008f6","l129|100|0025d8","l127|100|001be3","l124|100|00d32b","l123|100|0017e7","l122|100|000000","l121|100|00807e","l116|100|003ec0","l114|100|003ac4","l112|100|00ae50","l111|100|00ce30","l132|100|00a05e","l94|100|00c13d","l97|100|0008f6")
	End With
	With .AddStep(0.40000, Null, Null)
		.Lights = Array("l141|100|0043bb","l79|100|00649a","l78|100|00c737","l77|100|005aa4","l72|100|00ea14","l23|100|00926c","l61|100|00a559","l18|100|00f10d","l17|100|00d12d","l139|100|0020de","l135|100|0032cc","l131|100|000000","l130|100|000000","l129|100|0005f9","l127|100|000000","l124|100|00b24c","l123|100|000000","l121|100|00609e","l116|100|001ee0","l114|100|0019e5","l112|100|008c72","l111|100|00ad51","l132|100|007f7f","l94|100|009f5f","l97|100|000000")
	End With
	With .AddStep(0.43333, Null, Null)
		.Lights = Array("l141|100|0022dc","l79|100|0042bc","l78|100|00a559","l77|100|0037c7","l72|100|00c836","l71|100|00e816","l23|100|00718d","l61|100|00847a","l19|100|00f00e","l18|100|00cf2f","l17|100|00af4f","l139|100|0000ff","l135|100|0010ee","l129|100|000000","l124|100|00906e","l121|100|003dc1","l116|100|000000","l114|100|000000","l112|100|006b93","l111|100|008b73","l132|100|005ea0","l94|100|007e80")
	End With
	With .AddStep(0.46667, Null, Null)
		.Lights = Array("l141|100|0000fe","l79|100|0020de","l78|100|00837b","l77|100|0016e8","l72|100|00a658","l71|100|00c737","l70|100|00e21c","l69|100|00fc02","l23|100|004eaf","l61|100|00629c","l19|100|00cf2f","l18|100|00ae50","l17|100|008d71","l139|100|000000","l135|100|000000","l125|100|00f905","l124|100|006e90","l121|100|001ce2","l112|100|0049b5","l111|100|006995","l132|100|003cc2","l94|100|005ca1")
	End With
	With .AddStep(0.50000, Null, Null)
		.Lights = Array("l141|100|000000","l79|100|0000ff","l78|100|00629c","l77|100|000000","l72|100|00847a","l71|100|00a45a","l70|100|00c03e","l69|100|00da24","l23|100|002dd1","l62|100|00e519","l61|100|0040be","l19|100|00ad51","l18|100|008b73","l17|100|006b93","l46|100|00f806","l137|100|00ed11","l125|100|00d826","l124|100|004bb2","l121|100|000000","l112|100|0027d7","l111|100|0047b7","l132|100|001ae4","l94|100|003ac4")
	End With
	With .AddStep(0.53333, Null, Null)
		.Lights = Array("l79|100|000000","l78|100|003fbf","l72|100|00629c","l71|100|00827c","l70|100|009d61","l69|100|00b846","l23|100|000bf3","l35|100|00eb13","l62|100|00c33b","l34|100|00e915","l61|100|001ee0","l19|100|008a74","l18|100|006995","l17|100|0048b5","l46|100|00d628","l47|100|00f707","l137|100|00cb33","l125|100|00b648","l124|100|002ad4","l112|100|0005f9","l111|100|0025d9","l132|100|000000","l94|100|0018e6")
	End With
	With .AddStep(0.56667, Null, Null)
		.Lights = Array("l78|100|001ee0","l72|100|0040be","l71|100|00619d","l70|100|007c82","l69|100|009668","l23|100|000000","l63|100|00f30b","l35|100|00ca34","l62|100|00a05e","l34|100|00c737","l61|100|000000","l19|100|006995","l18|100|0047b7","l17|100|0027d7","l30|100|00ff00","l26|100|00fd02","l25|100|00e01e","l46|100|00b44a","l47|100|00d529","l64|100|00f905","l137|100|00aa54","l128|100|00f30b","l126|100|00eb13","l125|100|00936b","l124|100|0008f6","l112|100|000000","l111|100|0004fb","l94|100|000000")
	End With
	With .AddStep(0.60000, Null, Null)
		.Lights = Array("l143|100|00f30b","l78|100|000000","l72|100|001fdf","l71|100|003fbf","l70|100|005ba3","l69|100|007589","l63|100|00d22c","l35|100|00a856","l62|100|007f7f","l34|100|00a559","l92|100|00e915","l19|100|0047b7","l18|100|0026d8","l17|100|0006f8","l31|100|00f608","l30|100|00de20","l28|100|00fd02","l26|100|00db23","l25|100|00bf3f","l46|100|00926c","l47|100|00b44a","l64|100|00d826","l95|100|00ff00","l137|100|008876","l128|100|00d22c","l126|100|00ca34","l125|100|00728c","l124|100|000000","l120|100|00fd02","l111|100|000000")
	End With
	With .AddStep(0.63333, Null, Null)
		.Lights = Array("l142|100|00e41a","l143|100|00d22c","l72|100|0000ff","l71|100|001ee0","l70|100|0039c5","l69|100|0054aa","l63|100|00b14d","l35|100|008777","l62|100|005f9f","l34|100|008479","l92|100|00c935","l60|100|00e519","l19|100|0026d8","l18|100|0005f9","l17|100|000000","l33|100|00f905","l32|100|00ff00","l31|100|00d628","l30|100|00bd41","l28|100|00dc22","l27|100|00e01e","l26|100|00bb43","l25|100|009d61","l46|100|00728c","l47|100|00936b","l64|100|00b747","l95|100|00df1f","l137|100|006797","l128|100|00b14d","l126|100|00a955","l125|100|0051ad","l120|100|00dc22","l110|100|00f707")
	End With
	With .AddStep(0.66667, Null, Null)
		.Lights = Array("l142|100|00c43a","l143|100|00b24c","l72|100|000000","l71|100|0000ff","l70|100|0019e5","l69|100|0033cb","l63|100|00906e","l35|100|006797","l62|100|003ec0","l34|100|006599","l92|100|00a955","l60|100|00c539","l19|100|0006f8","l18|100|000000","l33|100|00d925","l32|100|00de20","l31|100|00b648","l30|100|009c62","l41|100|00fe01","l40|100|00f707","l39|100|00f00e","l38|100|00ea14","l37|100|00e41a","l28|100|00bc42","l27|100|00c13d","l26|100|009a64","l25|100|007d80","l46|100|0051ad","l47|100|00738b","l64|100|009668","l95|100|00bf3f","l137|100|0046b8","l128|100|00906e","l126|100|008876","l125|100|0031cd","l120|100|00bc42","l110|100|00d727")
	End With
	With .AddStep(0.70000, Null, Null)
		.Lights = Array("l142|100|00a45a","l143|100|00926b","l71|100|000000","l70|100|000000","l69|100|0014ea","l63|100|00718d","l35|100|0047b7","l62|100|001fdf","l34|100|0045b9","l92|100|008975","l60|100|00a558","l19|100|000000","l49|100|00fa04","l65|100|00f40a","l33|100|00ba43","l32|100|00bf3f","l31|100|009668","l30|100|007d81","l41|100|00de20","l40|100|00d826","l39|100|00d12d","l38|100|00cb33","l37|100|00c539","l28|100|009c62","l27|100|00a15d","l26|100|007b83","l25|100|005f9f","l46|100|0032cc","l47|100|0054aa","l64|100|007787","l95|100|009f5e","l137|100|0027d7","l128|100|00718d","l126|100|006995","l125|100|0012ec","l120|100|009c62","l119|100|00f40a","l110|100|00b846","l109|100|00ff00","l98|100|00f10d")
	End With
	With .AddStep(0.73333, Null, Null)
		.Lights = Array("l142|100|008677","l143|100|007589","l69|100|000000","l63|100|0053ab","l35|100|0029d5","l62|100|0001fd","l34|100|0027d7","l92|100|006b93","l60|100|008876","l49|100|00dd21","l65|100|00d727","l33|100|009c62","l32|100|00a05e","l31|100|007886","l30|100|00609e","l41|100|00c03e","l40|100|00ba44","l39|100|00b34b","l38|100|00ad51","l37|100|00a757","l28|100|007e80","l27|100|00837b","l26|100|005da1","l25|100|0040be","l46|100|0014ea","l47|100|0035c9","l64|100|005aa4","l95|100|00827c","l137|100|0009f4","l128|100|0053ab","l126|100|004ab4","l125|100|000000","l120|100|007e80","l119|100|00d727","l110|100|009965","l109|100|00e01e","l93|100|00ed11","l98|100|00d32b")
	End With
	With .AddStep(0.76667, Null, Null)
		.Lights = Array("l142|100|006a94","l143|100|0058a5","l63|100|0036c8","l35|100|000df1","l62|100|000000","l34|100|000bf3","l92|100|004eb0","l60|100|006b93","l49|100|00c03e","l59|100|00ff00","l65|100|00ba44","l33|100|007f7f","l32|100|00847a","l31|100|005ca2","l30|100|0042bc","l41|100|00a35b","l40|100|009d61","l39|100|009668","l38|100|00906e","l37|100|008a74","l28|100|00629c","l27|100|006797","l26|100|0040be","l25|100|0023da","l46|100|000000","l47|100|0019e5","l64|100|003cc2","l95|100|006598","l137|100|000000","l128|100|0036c8","l126|100|002ed0","l120|100|00629c","l119|100|00ba44","l110|100|007d81","l109|100|00c43a","l117|100|00ff00","l93|100|00d12d","l98|100|00b747")
	End With
	With .AddStep(0.80000, Null, Null)
		.Lights = Array("l142|100|004faf","l143|100|003dc1","l06|100|00fe01","l04|100|00f806","l03|100|00f40a","l02|100|00f00e","l01|100|00ea14","l63|100|001ce2","l35|100|000000","l34|100|000000","l92|100|0034ca","l60|100|0050ae","l49|100|00a559","l59|100|00e31b","l65|100|009f5f","l33|100|006599","l32|100|006a94","l31|100|0040bd","l30|100|0028d6","l41|100|008876","l40|100|00827b","l39|100|007b83","l38|100|007588","l37|100|00708e","l28|100|0046b8","l27|100|004bb3","l26|100|0025d9","l25|100|0009f5","l48|100|00f608","l47|100|0000ff","l64|100|0022dc","l95|100|004ab4","l128|100|001ce2","l126|100|0014ea","l120|100|0046b8","l119|100|009f5f","l110|100|00639b","l109|100|00a955","l117|100|00e31b","l93|100|00b648","l98|100|009b63")
	End With
	With .AddStep(0.83333, Null, Null)
		.Lights = Array("l142|100|0036c8","l143|100|0025d9","l05|100|00f40a","l07|100|00f905","l08|100|00ff00","l06|100|00e519","l04|100|00e01e","l03|100|00db23","l02|100|00d826","l01|100|00d22c","l63|100|0004fa","l92|100|001be3","l60|100|0038c6","l49|100|008c71","l59|100|00cb33","l65|100|008777","l33|100|004cb2","l32|100|0050ad","l31|100|0028d6","l30|100|000fef","l41|100|00708e","l40|100|006a94","l39|100|00639b","l38|100|005da1","l37|100|0057a7","l28|100|002ed0","l27|100|0033cb","l26|100|000df1","l25|100|000000","l48|100|00de20","l47|100|000000","l64|100|000af4","l95|100|0032cc","l128|100|0004fa","l126|100|000000","l120|100|002ed0","l119|100|008777","l110|100|0049b5","l109|100|00906e","l117|100|00cb33","l93|100|009d61","l98|100|00837b")
	End With
	With .AddStep(0.86667, Null, Null)
		.Lights = Array("l142|100|0021dd","l143|100|000fef","l05|100|00df1f","l07|100|00e31b","l08|100|00e915","l09|100|00ed11","l10|100|00f10c","l06|100|00cf2f","l04|100|00ca33","l03|100|00c638","l02|100|00c23c","l01|100|00bc42","l63|100|000000","l92|100|0006f8","l60|100|0022dc","l49|100|007787","l59|100|00b549","l65|100|00718d","l33|100|0036c8","l32|100|003bc3","l31|100|0013eb","l30|100|000000","l41|100|005aa3","l40|100|0054aa","l39|100|004cb1","l38|100|0047b7","l37|100|0041bd","l28|100|0018e5","l27|100|001de1","l26|100|000000","l24|100|00f30b","l48|100|00c836","l64|100|000000","l95|100|001ce2","l128|100|000000","l120|100|0018e5","l119|100|00718d","l110|100|0034ca","l109|100|007a84","l118|100|00f509","l117|100|00b549","l93|100|008777","l98|100|006d91")
	End With
	With .AddStep(0.90000, Null, Null)
		.Lights = Array("l142|100|000eef","l143|100|000000","l11|100|00f00e","l12|100|00f40a","l13|100|00f806","l14|100|00fd01","l05|100|00cc32","l07|100|00d12d","l08|100|00d727","l09|100|00da23","l10|100|00df1f","l06|100|00bd41","l04|100|00b846","l03|100|00b34a","l02|100|00b04e","l01|100|00aa54","l92|100|000000","l60|100|0010ee","l49|100|006499","l59|100|00a25c","l65|100|005f9f","l33|100|0024da","l32|100|0028d5","l31|100|0000ff","l29|100|00ff00","l41|100|0047b7","l40|100|0041bd","l39|100|003ac4","l38|100|0034ca","l37|100|002ed0","l28|100|0006f8","l27|100|000bf3","l24|100|00e01e","l48|100|00b648","l95|100|000af4","l120|100|0006f8","l119|100|005f9f","l110|100|0021dd","l109|100|006896","l108|100|00f806","l118|100|00e31b","l117|100|00a25c","l93|100|007589","l98|100|005ba3")
	End With
	With .AddStep(0.93333, Null, Null)
		.Lights = Array("l142|100|0000ff","l11|100|00e21c","l12|100|00e618","l13|100|00ea14","l14|100|00ef0f","l15|100|00f30b","l05|100|00be40","l07|100|00c33b","l08|100|00c935","l09|100|00cc32","l10|100|00d12d","l06|100|00af4f","l04|100|00aa54","l03|100|00a45a","l02|100|00a15d","l01|100|009b63","l60|100|0001fd","l49|100|0056a8","l59|100|00946a","l65|100|004faf","l33|100|0016e8","l32|100|001ae4","l31|100|000000","l29|100|00f10d","l41|100|0039c5","l40|100|0033cb","l39|100|002cd2","l38|100|0026d8","l37|100|0020de","l28|100|000000","l27|100|000000","l24|100|00d22c","l48|100|00a757","l95|100|000000","l120|100|000000","l119|100|004faf","l110|100|0013eb","l109|100|005aa4","l108|100|00ea14","l118|100|00d529","l117|100|00946a","l93|100|006797","l98|100|004cb2")
	End With
	With .AddStep(0.96667, Null, Null)
		.Lights = Array("l142|100|000000","l11|100|00d925","l12|100|00de20","l13|100|00e11d","l14|100|00e618","l15|100|00eb13","l05|100|00b549","l07|100|00ba44","l08|100|00c03e","l09|100|00c43a","l10|100|00c836","l06|100|00a559","l04|100|00a05e","l03|100|009c62","l02|100|009866","l01|100|00926c","l60|100|000000","l49|100|004db1","l59|100|008b73","l65|100|0047b7","l33|100|000df1","l32|100|0012ec","l29|100|00e816","l41|100|0030ce","l40|100|002ad4","l39|100|0023db","l38|100|001de1","l37|100|0017e6","l24|100|00c934","l48|100|009e60","l119|100|0047b7","l110|100|000af3","l109|100|0050ae","l108|100|00e11d","l118|100|00cc32","l117|100|008b73","l93|100|005ea0","l98|100|0043bb")
	End With
	With .AddStep(1.00000, Null, Null)
		.Lights = Array("l11|100|00d727","l12|100|00dc22","l13|100|00df1f","l14|100|00e41a","l15|100|00e915","l05|100|00b44a","l07|100|00b846","l08|100|00be40","l09|100|00c23c","l10|100|00c638","l06|100|00a35b","l04|100|009e60","l03|100|009a64","l02|100|009668","l01|100|00906e","l49|100|004ab4","l59|100|008975","l65|100|0045b9","l33|100|000bf3","l32|100|0010ee","l29|100|00e618","l41|100|002ed0","l40|100|0028d6","l39|100|0021dd","l38|100|001be3","l37|100|0015e9","l24|100|00c737","l48|100|009c62","l119|100|0045b9","l110|100|0008f6","l109|100|004eb0","l108|100|00df1f","l118|100|00ca34","l117|100|008975","l93|100|005ca2","l98|100|0041bd")
	End With
	With .AddStep(1.03333, Null, Null)
		.Lights = Array("l11|100|00db23","l12|100|00df1f","l13|100|00e21c","l14|100|00e618","l15|100|00ea14","l05|100|00b747","l07|100|00bb43","l08|100|00c13d","l09|100|00c33a","l10|100|00c836","l06|100|00a45a","l04|100|00a05e","l03|100|009c62","l02|100|009965","l01|100|00946a","l49|100|0047b7","l59|100|008579","l65|100|0048b6","l33|100|000fef","l29|100|00e21c","l40|100|0027d6","l39|100|0020de","l38|100|0019e5","l37|100|0012ec","l24|100|00c33b","l48|100|009767","l119|100|004bb3","l110|100|0002fd","l109|100|0049b4","l108|100|00d826","l118|100|00cf2f","l117|100|008e70","l93|100|005ea0","l98|100|003fbf")
	End With
	With .AddStep(1.06667, Null, Null)
		.Lights = Array("l142|100|0006f8","l11|100|00e21c","l12|100|00e519","l13|100|00e717","l14|100|00ea14","l15|100|00ec12","l05|100|00be40","l07|100|00c13d","l08|100|00c539","l09|100|00c737","l10|100|00ca34","l06|100|00a657","l04|100|00a35b","l03|100|00a05e","l02|100|009f5f","l01|100|009b63","l49|100|0041bd","l59|100|007d80","l65|100|004db0","l33|100|0016e8","l32|100|0012ec","l29|100|00da24","l40|100|0026d8","l39|100|001de1","l38|100|0014ea","l37|100|000cf2","l24|100|00bb43","l48|100|008e70","l120|100|0000ff","l119|100|0057a7","l110|100|000000","l109|100|0041bd","l108|100|00cb33","l118|100|00d727","l117|100|009668","l93|100|00619d","l98|100|003bc3")
	End With
	With .AddStep(1.10000, Null, Null)
		.Lights = Array("l142|100|0015e9","l11|100|00eb13","l12|100|00ed11","l13|100|00ed11","l14|100|00ee10","l15|100|00ef0f","l05|100|00c737","l07|100|00c935","l08|100|00cb33","l09|100|00cb33","l10|100|00cd31","l06|100|00aa54","l04|100|00a856","l03|100|00a658","l02|100|00a657","l01|100|00a45a","l49|100|0039c5","l59|100|00738a","l65|100|0056a8","l33|100|0020de","l32|100|0014ea","l29|100|00d02e","l41|100|002ecf","l40|100|0024da","l39|100|0019e5","l38|100|000fef","l37|100|0005f9","l28|100|0004fa","l24|100|00b04e","l48|100|00837b","l81|100|00ff00","l80|100|00f707","l120|100|000fef","l119|100|006598","l109|100|0036c8","l108|100|00ba44","l118|100|00e31b","l117|100|00a15d","l93|100|006698","l98|100|0037c7")
	End With
	With .AddStep(1.13333, Null, Null)
		.Lights = Array("l142|100|0027d7","l11|100|00f608","l12|100|00f509","l13|100|00f30a","l14|100|00f30b","l15|100|00f20c","l05|100|00d22c","l07|100|00d12d","l08|100|00d22c","l09|100|00d02e","l10|100|00cf2f","l06|100|00ad51","l04|100|00ad51","l03|100|00ae50","l02|100|00b04e","l01|100|00b04e","l49|100|0030ce","l59|100|006896","l65|100|00609e","l33|100|002cd1","l32|100|0018e6","l29|100|00c23c","l41|100|002fcf","l40|100|0023db","l39|100|0015e9","l38|100|0009f5","l37|100|000000","l28|100|0012ec","l24|100|00a15d","l43|100|00f806","l48|100|00748a","l42|100|00ee10","l81|100|00f707","l80|100|00eb13","l136|100|00f509","l120|100|0023db","l119|100|007787","l109|100|0029d5","l108|100|00a35b","l118|100|00f00e","l117|100|00b04e","l93|100|006b93","l98|100|0031cc")
	End With
	With .AddStep(1.16667, Null, Null)
		.Lights = Array("l142|100|003cc2","l11|100|000000","l12|100|00ff00","l13|100|00fa04","l14|100|00f707","l15|100|00f409","l05|100|00dd21","l07|100|00db23","l08|100|00d925","l09|100|00d529","l10|100|00d22c","l06|100|00b04e","l04|100|00b24c","l03|100|00b549","l02|100|00b945","l01|100|00bb42","l63|100|0000ff","l92|100|0004fa","l49|100|0026d8","l59|100|005aa4","l65|100|006b93","l33|100|003bc3","l32|100|001de1","l29|100|00b34b","l41|100|0031cd","l40|100|0022dc","l39|100|0012ec","l38|100|0003fb","l28|100|0022dc","l27|100|0003fb","l24|100|00916d","l43|100|00e21c","l48|100|006599","l42|100|00d42a","l82|100|00fe01","l81|100|00ec11","l80|100|00dc22","l136|100|00e31b","l120|100|0039c4","l119|100|008c72","l109|100|001be3","l108|100|008b73","l118|100|00ff00","l117|100|00be40","l93|100|00718d","l98|100|002cd2")
	End With
	With .AddStep(1.20000, Null, Null)
		.Lights = Array("l142|100|0055a9","l12|100|000000","l13|100|000000","l14|100|00fc03","l15|100|00f608","l05|100|00e915","l07|100|00e41a","l08|100|00e01e","l09|100|00d924","l10|100|00d42a","l06|100|00b24c","l04|100|00b747","l03|100|00bc42","l02|100|00c33b","l01|100|00c836","l63|100|0015e9","l92|100|0014ea","l49|100|001ce2","l59|100|004bb3","l65|100|007787","l33|100|004cb2","l32|100|0022db","l29|100|00a05d","l41|100|0033cb","l39|100|000fef","l38|100|000000","l28|100|0035c9","l27|100|000bf3","l24|100|00807e","l43|100|00ca34","l50|100|00ff00","l48|100|0053ab","l42|100|00b846","l82|100|00f508","l81|100|00e01e","l80|100|00cb33","l136|100|00ce30","l120|100|0054aa","l119|100|00a25c","l109|100|000df1","l108|100|00708e","l118|100|000000","l117|100|00ce30","l93|100|007886","l98|100|0027d7")
	End With
	With .AddStep(1.23333, Null, Null)
		.Lights = Array("l142|100|006f8f","l14|100|00ff00","l15|100|00f707","l05|100|00f40a","l07|100|00ed11","l08|100|00e618","l09|100|00dd20","l10|100|00d628","l06|100|00b549","l04|100|00bc42","l03|100|00c43a","l02|100|00cd31","l01|100|00d42a","l63|100|002dd1","l92|100|0026d8","l49|100|0012ec","l59|100|003cc2","l65|100|008479","l33|100|005f9f","l32|100|002ad4","l29|100|008e70","l41|100|0036c8","l39|100|000df1","l28|100|0049b5","l27|100|0015e9","l24|100|006e90","l43|100|00af4f","l50|100|00ec12","l48|100|0040bd","l42|100|009866","l82|100|00ec12","l81|100|00d22c","l80|100|00b846","l136|100|00b747","l120|100|00708e","l119|100|00ba44","l109|100|0000ff","l108|100|0054aa","l107|100|00ec12","l117|100|00dd20","l93|100|00807e","l98|100|0022dc")
	End With
	With .AddStep(1.26667, Null, Null)
		.Lights = Array("l142|100|008b73","l14|100|000000","l15|100|00f608","l05|100|00ff00","l07|100|00f509","l08|100|00ec12","l09|100|00e11d","l06|100|00b747","l04|100|00c03e","l03|100|00cb33","l02|100|00d727","l01|100|00e11d","l63|100|0048b6","l92|100|0039c5","l49|100|0009f5","l59|100|002ed0","l65|100|00926c","l33|100|00738b","l32|100|0032cc","l29|100|007a84","l41|100|003ac4","l40|100|0024da","l39|100|000cf2","l28|100|00619d","l27|100|0020de","l24|100|005aa4","l43|100|00926c","l50|100|00d628","l51|100|00f509","l48|100|002ed0","l64|100|0008f6","l42|100|007787","l83|100|00ff00","l82|100|00e11d","l81|100|00c23c","l80|100|00a35b","l136|100|009e60","l120|100|008d71","l119|100|00d22c","l109|100|000000","l108|100|0035c8","l107|100|00cf2f","l106|100|00ff00","l117|100|00ed11","l93|100|008876","l98|100|001fdf")
	End With
	With .AddStep(1.30000, Null, Null)
		.Lights = Array("l142|100|00a757","l69|100|0006f8","l15|100|00f40a","l05|100|000000","l07|100|00fd02","l08|100|00f00e","l09|100|00e31b","l06|100|00b846","l04|100|00c43a","l03|100|00d12d","l02|100|00df1f","l01|100|00ec12","l63|100|006599","l35|100|0002fd","l92|100|004faf","l49|100|0001fe","l59|100|001fdf","l65|100|00a05e","l33|100|008876","l32|100|003cc2","l29|100|006599","l41|100|003fbf","l40|100|0027d7","l28|100|007886","l27|100|002dd1","l24|100|0046b8","l43|100|00748a","l50|100|00be40","l51|100|00e21c","l52|100|00f20c","l48|100|001ce1","l64|100|001ae4","l42|100|0055a9","l83|100|00f707","l82|100|00d529","l81|100|00b04e","l80|100|008d71","l136|100|00847a","l120|100|00ad51","l119|100|00ea14","l108|100|0018e6","l107|100|00b04e","l106|100|00de20","l117|100|00fc02","l93|100|00906e","l98|100|001ce2")
	End With
	With .AddStep(1.33333, Null, Null)
		.Lights = Array("l142|100|00c539","l70|100|0021dd","l69|100|0028d6","l14|100|00ff00","l15|100|00f00e","l07|100|000000","l08|100|00f40a","l09|100|00e41a","l10|100|00d42a","l06|100|00b945","l04|100|00c737","l03|100|00d627","l02|100|00e717","l01|100|00f707","l63|100|00827c","l35|100|001ae4","l92|100|006698","l49|100|000000","l59|100|0012ec","l65|100|00b04e","l33|100|009d61","l32|100|0047b7","l31|100|0002fd","l29|100|004faf","l41|100|0045b9","l40|100|002bd3","l39|100|000ef0","l28|100|00916d","l27|100|003bc3","l24|100|0032cc","l43|100|0055a9","l50|100|00a35b","l51|100|00cd31","l52|100|00d727","l48|100|000bf3","l64|100|002dd1","l42|100|0032cc","l83|100|00ed11","l82|100|00c737","l81|100|009d61","l80|100|007688","l136|100|006995","l120|100|00cb33","l119|100|000000","l108|100|000000","l107|100|008f6f","l106|100|00bb43","l117|100|000000","l93|100|009866","l98|100|001ae4")
	End With
	With .AddStep(1.36667, Null, Null)
		.Lights = Array("l142|100|00e11d","l70|100|0048b6","l69|100|004cb2","l14|100|00fc02","l15|100|00ea14","l08|100|00f509","l09|100|00e31b","l10|100|00d22c","l04|100|00c934","l03|100|00db23","l02|100|00ee10","l01|100|00ff00","l63|100|00a05e","l53|100|00f00e","l35|100|0034c9","l92|100|007e80","l59|100|0005f9","l65|100|00bd40","l33|100|00b34b","l32|100|0054aa","l31|100|000df1","l29|100|003ac4","l41|100|004cb2","l40|100|0030ce","l39|100|0012ec","l28|100|00aa53","l27|100|004bb3","l24|100|001fdf","l43|100|0035c9","l50|100|008876","l51|100|00b648","l52|100|00bb43","l20|100|00f608","l56|100|00f806","l54|100|00f30b","l48|100|000000","l64|100|0041bd","l42|100|000fef","l83|100|00e21c","l82|100|00b746","l81|100|008975","l80|100|005f9f","l136|100|004db1","l120|100|00ea14","l107|100|006d91","l106|100|009569","l93|100|00a05e")
	End With
	With .AddStep(1.40000, Null, Null)
		.Lights = Array("l142|100|00fe01","l70|100|00728c","l69|100|00728c","l14|100|00f707","l15|100|00e31b","l09|100|00e11d","l10|100|00ce30","l06|100|00b846","l04|100|00cb33","l03|100|00de20","l02|100|00f30b","l01|100|000000","l63|100|00bf3f","l53|100|00d628","l35|100|0050ad","l62|100|000bf3","l34|100|0012ec","l92|100|009668","l59|100|000000","l65|100|00cb33","l33|100|00c836","l32|100|00619d","l31|100|001ae4","l29|100|0025d9","l41|100|0055a9","l40|100|0036c8","l39|100|0017e7","l28|100|00c33b","l27|100|005da1","l24|100|000df1","l43|100|0016e8","l50|100|006d91","l51|100|009e60","l52|100|009c62","l20|100|00e31a","l57|100|00f806","l56|100|00d42a","l55|100|00f904","l54|100|00d42a","l64|100|0058a5","l42|100|000000","l83|100|00d529","l82|100|00a658","l81|100|007589","l80|100|0047b7","l136|100|0031cc","l120|100|000000","l107|100|004ab4","l106|100|00708e","l93|100|00a955","l98|100|001be3")
	End With
	With .AddStep(1.43333, Null, Null)
		.Lights = Array("l142|100|000000","l70|100|009c62","l69|100|009866","l14|100|00f00e","l15|100|00db23","l08|100|00f40a","l09|100|00de20","l10|100|00c935","l06|100|00b648","l03|100|00e01e","l02|100|00f707","l63|100|00dd21","l53|100|00ba44","l35|100|006e90","l62|100|002ad4","l34|100|002cd2","l92|100|00af4f","l65|100|00d727","l33|100|00dc22","l32|100|006f8f","l31|100|0029d4","l29|100|0012ec","l41|100|005ea0","l40|100|003ec0","l39|100|001ee0","l38|100|0000ff","l28|100|00db23","l27|100|006e90","l26|100|000fef","l24|100|000000","l43|100|000000","l50|100|0050ae","l51|100|008579","l52|100|007d81","l22|100|00f707","l20|100|00cf2f","l58|100|00fd02","l57|100|00d529","l56|100|00ae50","l55|100|00db23","l54|100|00b34b","l64|100|006f8f","l84|100|00f905","l83|100|00c638","l82|100|009569","l81|100|00619d","l80|100|0030ce","l136|100|0017e7","l107|100|0029d5","l106|100|0049b5","l93|100|00b04e","l98|100|001de0")
	End With
	With .AddStep(1.46667, Null, Null)
		.Lights = Array("l70|100|00c638","l69|100|00bf3f","l13|100|00ff00","l14|100|00e717","l15|100|00d02d","l08|100|00f00e","l09|100|00d925","l10|100|00c23c","l06|100|00b44a","l04|100|00ca34","l02|100|00f905","l63|100|00fa04","l53|100|009c62","l35|100|008c72","l62|100|004ab4","l34|100|0047b7","l92|100|00c638","l60|100|0002fc","l65|100|00e21c","l33|100|00ef0f","l32|100|007d81","l31|100|003ac4","l29|100|0000ff","l41|100|006797","l40|100|0047b7","l39|100|0026d8","l38|100|0007f7","l28|100|00f20c","l27|100|00807e","l26|100|0022dc","l50|100|0035c9","l51|100|006c92","l52|100|005ea0","l21|100|00ef0f","l22|100|00dd21","l20|100|00b945","l58|100|00d826","l57|100|00af4f","l56|100|008678","l55|100|00bb43","l54|100|00906e","l47|100|0001fd","l64|100|008777","l84|100|00eb13","l83|100|00b648","l82|100|00837b","l81|100|004cb2","l80|100|001be3","l136|100|0000ff","l107|100|0008f6","l106|100|0024da","l93|100|00b747","l98|100|0022dc")
	End With
	With .AddStep(1.50000, Null, Null)
		.Lights = Array("l71|100|00fa04","l70|100|00ef0f","l69|100|00e419","l13|100|00f509","l14|100|00dd21","l15|100|00c539","l08|100|00eb13","l09|100|00d32b","l10|100|00bb43","l06|100|00b14d","l04|100|00c737","l03|100|00df1f","l63|100|000000","l53|100|007e7f","l35|100|00ab53","l62|100|006c92","l34|100|00649a","l92|100|00dd21","l60|100|0011ed","l19|100|0020de","l18|100|001ce2","l65|100|00ec12","l33|100|000000","l32|100|008b73","l31|100|004cb2","l29|100|000000","l41|100|00718d","l40|100|0050ad","l39|100|0030ce","l38|100|0011ed","l28|100|000000","l27|100|00926c","l26|100|0036c7","l50|100|001be3","l51|100|0052ab","l52|100|003ec0","l21|100|00d925","l22|100|00c23c","l20|100|00a15d","l58|100|00b34b","l57|100|008876","l56|100|005f9f","l55|100|009866","l54|100|006e90","l47|100|001ae4","l64|100|009e60","l84|100|00db23","l83|100|00a459","l82|100|00718d","l81|100|0039c5","l80|100|0006f8","l136|100|000000","l145|100|00ff00","l107|100|000000","l106|100|0000ff","l93|100|00bd41","l98|100|0027d7")
	End With
	With .AddStep(1.53333, Null, Null)
		.Lights = Array("l71|100|000000","l70|100|000000","l69|100|000000","l13|100|00ea14","l14|100|00d12d","l15|100|00b945","l07|100|00fe01","l08|100|00e41a","l09|100|00cb33","l10|100|00b24b","l06|100|00ac51","l04|100|00c43a","l03|100|00dd21","l02|100|00f707","l53|100|00619d","l35|100|00c836","l62|100|008e70","l34|100|00817d","l92|100|00f20c","l60|100|0021dd","l19|100|0048b6","l18|100|0048b6","l65|100|00f40a","l32|100|009965","l31|100|005f9f","l41|100|007a83","l40|100|005ca2","l39|100|003bc3","l38|100|001ce2","l27|100|00a45a","l26|100|004cb2","l50|100|0002fc","l51|100|003ac4","l52|100|0020de","l21|100|00c23c","l22|100|00a45a","l20|100|008a74","l58|100|008c72","l57|100|00619d","l56|100|0037c7","l55|100|007688","l54|100|004bb3","l47|100|0034ca","l64|100|00b648","l84|100|00ca34","l83|100|00936b","l82|100|005f9f","l81|100|0027d7","l80|100|000000","l137|100|0018e6","l145|100|00ed11","l128|100|0005f9","l106|100|000000","l105|100|00f607","l104|100|00eb13","l101|100|00ef0f","l93|100|00c13d","l98|100|002ed0")
	End With
	With .AddStep(1.56667, Null, Null)
		.Lights = Array("l12|100|00f707","l13|100|00de20","l14|100|00c43a","l15|100|00ab53","l07|100|00f509","l08|100|00dc22","l09|100|00c33b","l10|100|00a955","l06|100|00a757","l04|100|00c03e","l03|100|00d925","l02|100|00f30b","l53|100|0042bc","l35|100|00e41a","l62|100|00b04e","l34|100|009e60","l92|100|000000","l60|100|0032cc","l19|100|00728c","l18|100|000000","l49|100|0001fd","l65|100|00fa04","l32|100|00a757","l31|100|00728c","l41|100|00847a","l40|100|006797","l39|100|0047b7","l38|100|0029d5","l37|100|000bf3","l27|100|00b648","l26|100|00649a","l50|100|000000","l51|100|0022db","l52|100|0003fb","l21|100|00aa54","l22|100|008777","l20|100|00728c","l58|100|006599","l57|100|003ac4","l56|100|0011ed","l55|100|0054aa","l54|100|0029d5","l47|100|0050ae","l64|100|00cc32","l84|100|00b846","l83|100|00817d","l82|100|004db1","l81|100|0016e8","l95|100|0000ff","l137|100|0039c5","l145|100|00d826","l128|100|001de0","l105|100|00da24","l104|100|00c935","l102|100|00e915","l101|100|000000","l93|100|00c539","l98|100|0036c8")
	End With
	With .AddStep(1.60000, Null, Null)
		.Lights = Array("l44|100|00e816","l12|100|00e816","l13|100|00cf2e","l14|100|00b648","l15|100|009c62","l07|100|00eb13","l08|100|00d22c","l09|100|00b945","l10|100|009f5f","l06|100|00a15d","l04|100|00ba44","l03|100|00d32a","l02|100|00ee10","l53|100|0026d8","l35|100|00ff00","l62|100|00d12d","l34|100|00bb43","l60|100|0045b9","l19|100|009b63","l49|100|000bf3","l65|100|00ff00","l32|100|00b549","l31|100|008579","l30|100|0010ee","l41|100|008e70","l40|100|00728c","l39|100|0055a9","l38|100|0037c7","l37|100|001be3","l27|100|00c638","l26|100|007a83","l25|100|000df1","l51|100|000df1","l52|100|000000","l21|100|00916d","l22|100|006995","l20|100|005ba3","l58|100|003ec0","l57|100|0015e9","l56|100|000000","l55|100|0032cc","l54|100|0009f5","l46|100|0008f6","l47|100|006d91","l64|100|00e11d","l84|100|00a45a","l83|100|00708e","l82|100|003cc2","l81|100|0006f8","l95|100|0014ea","l137|100|005ca2","l145|100|00c23c","l128|100|0037c7","l105|100|00bd41","l104|100|00a559","l102|100|00be40","l93|100|00c836","l98|100|003fbf")
	End With
	With .AddStep(1.63333, Null, Null)
		.Lights = Array("l44|100|00cf2f","l11|100|00f20c","l12|100|00d925","l13|100|00c03e","l14|100|00a757","l15|100|008e70","l05|100|00f806","l07|100|00e01e","l08|100|00c737","l09|100|00ae50","l10|100|009469","l06|100|009b63","l04|100|00b44a","l03|100|00cd31","l02|100|00e618","l53|100|000cf2","l35|100|000000","l62|100|00f00e","l34|100|00d628","l60|100|0059a5","l19|100|00c43a","l18|100|00d02e","l49|100|0016e8","l65|100|000000","l32|100|00c03d","l31|100|009866","l30|100|002ad4","l41|100|009866","l40|100|007d81","l39|100|00629c","l38|100|0047b7","l37|100|002cd2","l27|100|00d529","l26|100|00916d","l25|100|002bd3","l45|100|00fa04","l51|100|000000","l21|100|007886","l22|100|004cb2","l20|100|0043bb","l58|100|0019e5","l57|100|000000","l55|100|0013eb","l54|100|000000","l46|100|002dd1","l47|100|008975","l64|100|00f409","l84|100|00916d","l83|100|005ea0","l82|100|002dd1","l81|100|000000","l95|100|002ad4","l137|100|007f7f","l145|100|00ac52","l133|100|00ff00","l128|100|0052ac","l105|100|009d61","l104|100|00827c","l102|100|00916d","l93|100|00ca34","l98|100|0049b4")
	End With
	With .AddStep(1.66667, Null, Null)
		.Lights = Array("l44|100|00b449","l11|100|00e01e","l12|100|00c836","l13|100|00b04e","l14|100|009767","l15|100|007f7f","l05|100|00eb13","l07|100|00d32a","l08|100|00bb43","l09|100|00a25c","l10|100|008a74","l06|100|009569","l04|100|00ad51","l03|100|00c539","l02|100|00de20","l01|100|00f806","l53|100|000000","l62|100|000000","l34|100|00ef0f","l60|100|006d91","l19|100|00ea14","l18|100|00fa04","l49|100|0022dc","l32|100|00cb32","l31|100|00ab53","l30|100|0045b8","l41|100|00a15d","l40|100|008975","l39|100|00708e","l38|100|0057a6","l37|100|003ec0","l27|100|00e31b","l26|100|00a856","l25|100|004bb3","l45|100|00dd21","l21|100|00609e","l22|100|0030ce","l20|100|002ed0","l58|100|000000","l55|100|000000","l46|100|0052ac","l47|100|00a559","l64|100|000000","l84|100|007e80","l83|100|004db1","l82|100|001fdf","l95|100|0041bd","l137|100|00a15d","l145|100|00936a","l133|100|00d826","l128|100|006d91","l126|100|0005f9","l105|100|007f7f","l104|100|005f9f","l102|100|006698","l98|100|0055a9")
	End With
	With .AddStep(1.70000, Null, Null)
		.Lights = Array("l143|100|0008f6","l44|100|009965","l11|100|00cd31","l12|100|00b648","l13|100|009f5f","l14|100|008876","l15|100|00718d","l05|100|00dc22","l07|100|00c638","l08|100|00ae50","l09|100|009767","l10|100|00807e","l06|100|008f6f","l04|100|00a559","l03|100|00bd41","l02|100|00d42a","l01|100|00ed11","l34|100|000000","l60|100|00817d","l19|100|000000","l18|100|000000","l49|100|0030ce","l32|100|00d529","l31|100|00bd41","l30|100|00629c","l41|100|00aa54","l40|100|00946a","l39|100|007e80","l38|100|006896","l37|100|0051ad","l27|100|00ef0f","l26|100|00be40","l25|100|006b93","l45|100|00c03e","l21|100|0047b7","l22|100|0016e8","l20|100|0019e4","l46|100|007886","l47|100|00c13d","l84|100|006a93","l83|100|003dc1","l82|100|0012ec","l95|100|0059a5","l137|100|00c33b","l145|100|007c82","l133|100|00b04e","l128|100|008876","l126|100|0028d6","l109|100|0008f6","l105|100|00619d","l104|100|003cc2","l102|100|003bc3","l101|100|000fef","l118|100|00ee10","l98|100|00619d")
	End With
	With .AddStep(1.73333, Null, Null)
		.Lights = Array("l143|100|0029d5","l44|100|007e80","l11|100|00ba44","l12|100|00a35b","l13|100|008e70","l14|100|007985","l15|100|00639b","l05|100|00cc32","l07|100|00b846","l08|100|00a05e","l09|100|008b73","l10|100|007688","l06|100|008876","l04|100|009d61","l03|100|00b34b","l02|100|00c935","l01|100|00e01d","l60|100|00946a","l49|100|003ec0","l59|100|0003fb","l65|100|00fe01","l32|100|00dd21","l31|100|00cd31","l30|100|007e80","l41|100|00b24c","l40|100|009e60","l39|100|008b73","l38|100|007886","l37|100|006599","l27|100|00f905","l26|100|00d22c","l25|100|008b73","l45|100|00a15d","l21|100|0031cd","l22|100|000000","l20|100|0007f7","l46|100|009c61","l47|100|00da24","l84|100|0058a6","l83|100|002ecf","l82|100|0008f6","l95|100|00718d","l137|100|00e21c","l145|100|006698","l133|100|008677","l128|100|00a25c","l126|100|004cb2","l110|100|000eef","l109|100|0019e5","l105|100|0043bb","l104|100|001be3","l102|100|0012ec","l101|100|000000","l118|100|00da24","l117|100|00fe01","l93|100|00c836","l98|100|006c92")
	End With
	With .AddStep(1.76667, Null, Null)
		.Lights = Array("l143|100|004ab4","l44|100|00649a","l11|100|00a559","l12|100|00916d","l13|100|007e80","l14|100|006a94","l15|100|0056a8","l05|100|00bc42","l07|100|00a855","l08|100|00936b","l09|100|00807e","l10|100|006c92","l06|100|00817d","l04|100|009569","l03|100|00aa54","l02|100|00be40","l01|100|00d32b","l60|100|00a757","l49|100|004db1","l59|100|000fef","l65|100|00f806","l32|100|00e41a","l31|100|00dc22","l30|100|009965","l41|100|00b945","l40|100|00a955","l39|100|009866","l38|100|008876","l37|100|007886","l27|100|000000","l26|100|00e41a","l25|100|00aa54","l45|100|00837b","l21|100|001ce2","l20|100|000000","l46|100|00c13d","l47|100|00f20c","l84|100|0045b8","l83|100|0021dd","l82|100|0000ff","l95|100|008876","l137|100|00ff00","l145|100|004faf","l133|100|005f9f","l128|100|00bc42","l126|100|00718d","l110|100|002dd1","l109|100|002cd2","l105|100|0027d7","l104|100|000000","l102|100|000000","l118|100|00c43a","l117|100|00ee10","l93|100|00c638","l98|100|007886")
	End With
	With .AddStep(1.80000, Null, Null)
		.Lights = Array("l143|100|006c92","l44|100|004bb3","l11|100|00926c","l12|100|007f7f","l13|100|006e90","l14|100|005ca2","l15|100|0049b5","l05|100|00ab52","l07|100|009965","l08|100|008678","l09|100|007589","l10|100|00629b","l06|100|007b83","l04|100|008d71","l03|100|009f5f","l02|100|00b24c","l01|100|00c539","l60|100|00b945","l49|100|005da1","l59|100|001ce2","l65|100|00f20c","l32|100|00ea14","l31|100|00e915","l30|100|00b44a","l41|100|00bf3f","l40|100|00b24c","l39|100|00a45a","l38|100|009767","l37|100|008a73","l26|100|00f509","l25|100|00c737","l45|100|006698","l21|100|0009f5","l46|100|00e21c","l47|100|000000","l84|100|0035c9","l83|100|0016e8","l82|100|000000","l95|100|009f5f","l137|100|000000","l145|100|003bc3","l133|100|0038c6","l128|100|00d42a","l126|100|00936b","l110|100|004bb3","l109|100|003fbf","l105|100|000ef0","l118|100|00af4f","l117|100|00de20","l93|100|00c23c","l98|100|00837b")
	End With
	With .AddStep(1.83333, Null, Null)
		.Lights = Array("l143|100|008b73","l44|100|0033cb","l11|100|007f7f","l12|100|006e90","l13|100|005f9f","l14|100|004eb0","l15|100|003ec0","l05|100|009a64","l07|100|008a73","l08|100|007985","l09|100|006a94","l10|100|005aa4","l06|100|007589","l04|100|008579","l03|100|009569","l02|100|00a459","l01|100|00b747","l60|100|00c935","l49|100|006c92","l59|100|002ad4","l65|100|00eb13","l32|100|00ed11","l31|100|00f40a","l30|100|00cd31","l41|100|00c43a","l40|100|00ba44","l39|100|00b14d","l38|100|00a658","l37|100|009c62","l26|100|000000","l25|100|00e21c","l45|100|004ab4","l21|100|000000","l48|100|000df1","l46|100|000000","l84|100|0026d8","l83|100|000cf2","l95|100|00b549","l145|100|0028d6","l133|100|0015e9","l128|100|00e915","l126|100|00b648","l110|100|006a94","l109|100|0052ac","l105|100|000000","l118|100|009965","l117|100|00cd31","l93|100|00be40","l98|100|008d70")
	End With
	With .AddStep(1.86667, Null, Null)
		.Lights = Array("l143|100|00aa54","l44|100|001ee0","l11|100|006c92","l12|100|005ea0","l13|100|0050ae","l14|100|0042bc","l15|100|0034ca","l05|100|008a74","l07|100|007d81","l08|100|006d91","l09|100|00609e","l10|100|0051ad","l06|100|006f8f","l04|100|007d81","l03|100|008b73","l02|100|009965","l01|100|00a955","l60|100|00d826","l49|100|007a84","l59|100|0038c6","l65|100|00e21b","l32|100|00f00e","l31|100|00ff00","l30|100|00e31b","l41|100|00c836","l40|100|00c13d","l39|100|00bb43","l38|100|00b44a","l37|100|00ae50","l25|100|00fb03","l45|100|0030ce","l48|100|001ce2","l84|100|0019e5","l83|100|0004fa","l95|100|00c935","l145|100|0017e7","l133|100|000000","l128|100|00fe01","l126|100|00d429","l110|100|008777","l109|100|006599","l118|100|00847a","l117|100|00bc42","l93|100|00ba44","l98|100|009766")
	End With
	With .AddStep(1.90000, Null, Null)
		.Lights = Array("l143|100|00c638","l44|100|000bf3","l11|100|005ca2","l12|100|004faf","l13|100|0044ba","l14|100|0038c6","l15|100|002bd2","l05|100|007b83","l07|100|006f8f","l08|100|00629c","l09|100|0057a7","l10|100|004ab4","l06|100|006a94","l04|100|007688","l03|100|00827c","l02|100|008d71","l01|100|009b63","l60|100|00e519","l49|100|008876","l59|100|0045b9","l65|100|00da24","l32|100|00f10d","l31|100|000000","l30|100|00f707","l41|100|00cb33","l40|100|00c737","l39|100|00c43a","l38|100|00c03e","l37|100|00bd41","l25|100|000000","l24|100|0007f7","l45|100|0019e5","l48|100|002bd3","l84|100|000ef0","l83|100|000000","l95|100|00da24","l145|100|0008f6","l128|100|000000","l126|100|00f10d","l119|100|00fc03","l110|100|00a15d","l109|100|007688","l118|100|00718d","l117|100|00ac52","l93|100|00b549","l98|100|00a15d")
	End With
	With .AddStep(1.93333, Null, Null)
		.Lights = Array("l143|100|00df1f","l44|100|000000","l11|100|004bb3","l12|100|0041bd","l13|100|0038c5","l14|100|002ed0","l15|100|0024da","l05|100|006d91","l07|100|00649a","l08|100|0058a6","l09|100|004faf","l10|100|0044ba","l06|100|006698","l04|100|006f8f","l03|100|007985","l02|100|00837b","l01|100|008e70","l60|100|00f00e","l49|100|00946a","l59|100|0053ab","l65|100|00d12d","l32|100|00f20c","l30|100|000000","l41|100|00cd31","l40|100|00cc32","l39|100|00cc32","l38|100|00cb33","l37|100|00ca34","l24|100|0014ea","l45|100|0005f9","l48|100|003ac4","l84|100|0004f9","l95|100|00ea14","l145|100|000000","l126|100|000000","l119|100|00ea14","l110|100|00bb43","l109|100|008777","l118|100|00609e","l117|100|009c62","l93|100|00b04e","l98|100|00aa54")
	End With
	With .AddStep(1.96667, Null, Null)
		.Lights = Array("l143|100|00f40a","l11|100|003ec0","l12|100|0036c8","l13|100|002fcf","l14|100|0027d7","l15|100|001fdf","l05|100|00619d","l07|100|0059a5","l08|100|004faf","l09|100|0048b6","l10|100|0040be","l06|100|00629c","l04|100|006a94","l03|100|00728c","l02|100|007985","l01|100|00837b","l60|100|00f905","l49|100|009f5f","l59|100|005f9f","l65|100|00c935","l32|100|00f10d","l29|100|0002fc","l41|100|00ce30","l40|100|00d02e","l39|100|00d22c","l38|100|00d42a","l37|100|00d529","l24|100|001fdf","l45|100|000000","l48|100|0047b7","l84|100|000000","l95|100|00f707","l119|100|00d925","l110|100|00d02e","l109|100|009569","l118|100|0050ae","l117|100|008e70","l93|100|00ac52","l98|100|00b14d")
	End With
	With .AddStep(2.00000, Null, Null)
		.Lights = Array("l143|100|000000","l11|100|0034ca","l12|100|002dd1","l13|100|0027d6","l14|100|0021dd","l15|100|001ae4","l05|100|0057a7","l07|100|0050ae","l08|100|0048b6","l09|100|0042bb","l10|100|003cc2","l06|100|005f9f","l04|100|006599","l03|100|006c92","l02|100|00718d","l01|100|007985","l60|100|00ff00","l49|100|00a955","l59|100|006995","l65|100|00c23c","l33|100|00ff00","l32|100|00f00e","l29|100|000cf2","l41|100|00cf2f","l40|100|00d32b","l39|100|00d727","l38|100|00db23","l37|100|00de20","l24|100|002ad4","l48|100|0054aa","l95|100|000000","l119|100|00cb33","l110|100|00e11d","l109|100|00a15d","l108|100|000bf3","l118|100|0043bb","l117|100|00837b","l93|100|00a757","l98|100|00b648")
	End With
	With .AddStep(2.03333, Null, Null)
		.Lights = Array("l11|100|002cd2","l12|100|0026d8","l13|100|0022dc","l14|100|001de1","l15|100|0017e7","l05|100|004faf","l07|100|004ab4","l08|100|0043bb","l09|100|003fbf","l10|100|0039c5","l06|100|005ca2","l04|100|00629c","l03|100|006797","l02|100|006b93","l01|100|00728c","l60|100|000000","l49|100|00af4e","l59|100|00708e","l65|100|00bd41","l33|100|00f707","l32|100|00ef0f","l29|100|0013eb","l41|100|00d02e","l40|100|00d529","l39|100|00db23","l38|100|00e01e","l37|100|00e519","l24|100|0032cc","l48|100|005da1","l119|100|00c03e","l110|100|00ee10","l109|100|00ab53","l108|100|0017e7","l118|100|003ac4","l117|100|007a84","l93|100|00a45a","l98|100|00ba43")
	End With
	With .AddStep(2.06667, Null, Null)
		.Lights = Array("l11|100|0028d6","l12|100|0023db","l13|100|001fdf","l14|100|001ae4","l15|100|0016e8","l05|100|004bb3","l07|100|0046b8","l08|100|0040be","l09|100|003dc1","l10|100|0038c6","l06|100|005ba3","l04|100|00609e","l03|100|006599","l02|100|006896","l01|100|006e90","l49|100|00b34b","l59|100|007589","l65|100|00ba44","l33|100|00f30a","l32|100|00ee10","l29|100|0017e7","l40|100|00d628","l39|100|00dd21","l38|100|00e21c","l37|100|00e816","l24|100|0036c8","l48|100|00629c","l119|100|00ba44","l110|100|00f509","l109|100|00af4f","l108|100|001ee0","l118|100|0035c9","l117|100|007589","l93|100|00a25c","l98|100|00bd41")
	End With
	With .AddStep(2.10000, Null, Null)
		.Lights = Array("l13|100|0020de","l14|100|001be3","l05|100|004cb2","l07|100|0047b7","l08|100|0041bd","l02|100|006995","l01|100|006f8f","l49|100|00b24c","l59|100|00748a","l33|100|00f40a","l32|100|00ef0f","l40|100|00d529","l39|100|00dc22","l24|100|0035c9","l48|100|00619d","l119|100|00bb43","l110|100|00f40a","l108|100|001de1","l117|100|007688","l98|100|00bc42")
	End With
	With .AddStep(2.13333, Null, Null)
		.Lights = Array("l11|100|002ed0","l12|100|0028d6","l13|100|0023db","l14|100|001ee0","l15|100|0018e6","l05|100|0051ad","l07|100|004bb3","l08|100|0044ba","l09|100|0040be","l10|100|003ac4","l06|100|005da1","l04|100|00639b","l03|100|006896","l02|100|006d91","l01|100|00748a","l49|100|00ae50","l59|100|006e8f","l65|100|00be40","l33|100|00f905","l29|100|0011ed","l41|100|00cf2f","l40|100|00d42a","l39|100|00da24","l38|100|00de20","l37|100|00e31b","l24|100|0030ce","l48|100|005aa3","l119|100|00c33b","l110|100|00eb13","l109|100|00a856","l108|100|0014ea","l118|100|003cc2","l117|100|007c82","l93|100|00a559","l98|100|00b945")
	End With
	With .AddStep(2.16667, Null, Null)
		.Lights = Array("l11|100|0037c7","l12|100|0030ce","l13|100|002ad4","l14|100|0023db","l15|100|001ce2","l05|100|005ba3","l07|100|0053ab","l08|100|004ab4","l09|100|0044ba","l10|100|003dc1","l06|100|00609e","l04|100|006797","l03|100|006e90","l02|100|00748a","l01|100|007c82","l60|100|00fe01","l49|100|00a559","l59|100|006698","l65|100|00c539","l33|100|000000","l32|100|00f00d","l29|100|0009f5","l40|100|00d22c","l39|100|00d628","l38|100|00d826","l37|100|00db23","l24|100|0026d8","l48|100|004faf","l95|100|00ff00","l119|100|00cf2e","l110|100|00dc22","l109|100|009d61","l108|100|0005f9","l118|100|0047b7","l117|100|008777","l93|100|00a955","l98|100|00b549")
	End With
	With .AddStep(2.20000, Null, Null)
		.Lights = Array("l143|100|00ec12","l11|100|0044ba","l12|100|003bc3","l13|100|0033cb","l14|100|002ad4","l15|100|0021dd","l05|100|006698","l07|100|005da0","l08|100|0053ab","l09|100|004bb3","l10|100|0042bc","l06|100|00639b","l04|100|006c92","l03|100|007589","l02|100|007d81","l01|100|008777","l60|100|00f509","l49|100|009b63","l59|100|005aa4","l65|100|00cd31","l32|100|00f10d","l29|100|0000ff","l41|100|00ce30","l40|100|00ce30","l39|100|00d02e","l38|100|00d02e","l37|100|00d12d","l24|100|001ae3","l48|100|0042bc","l84|100|0000ff","l95|100|00f20c","l119|100|00e01e","l110|100|00c737","l109|100|008f6f","l108|100|000000","l118|100|0057a7","l117|100|00946a","l93|100|00ae50","l98|100|00ae50")
	End With
	With .AddStep(2.23333, Null, Null)
		.Lights = Array("l143|100|00d32b","l44|100|0003fc","l11|100|0054aa","l12|100|0048b6","l13|100|003ec0","l14|100|0033cb","l15|100|0028d6","l05|100|00748a","l07|100|006a94","l08|100|005da1","l09|100|0053ab","l10|100|0047b7","l06|100|006896","l04|100|00738b","l03|100|007e80","l02|100|008876","l01|100|00946a","l60|100|00ea14","l49|100|008e70","l59|100|004cb2","l65|100|00d628","l32|100|00f10c","l30|100|00ff00","l29|100|000000","l41|100|00cc32","l40|100|00c934","l39|100|00c836","l38|100|00c638","l37|100|00c33b","l24|100|000df1","l45|100|000fef","l48|100|0032cb","l84|100|0009f5","l95|100|00e21c","l145|100|0001fe","l126|100|00fe01","l119|100|00f30b","l110|100|00af4f","l109|100|007f7f","l118|100|006995","l117|100|00a45a","l93|100|00b34b","l98|100|00a559")
	End With
	With .AddStep(2.26667, Null, Null)
		.Lights = Array("l143|100|00b548","l44|100|0016e8","l11|100|006698","l12|100|0058a6","l13|100|004bb3","l14|100|003ec0","l15|100|0030cd","l05|100|00847a","l07|100|007787","l08|100|006995","l09|100|005da1","l10|100|004eb0","l06|100|006d91","l04|100|007a84","l03|100|008777","l02|100|00946a","l01|100|00a35b","l60|100|00dd21","l49|100|00807e","l59|100|003dc1","l65|100|00df1f","l32|100|00f10d","l30|100|00eb13","l41|100|00c935","l40|100|00c33b","l39|100|00bf3f","l38|100|00b945","l37|100|00b44a","l24|100|0000ff","l45|100|0027d7","l48|100|0022dc","l84|100|0015e9","l83|100|0001fd","l95|100|00d02e","l145|100|0010ee","l126|100|00e01e","l119|100|000000","l110|100|00916c","l109|100|006c92","l118|100|007d81","l117|100|00b648","l93|100|00b846","l98|100|009b63")
	End With
	With .AddStep(2.30000, Null, Null)
		.Lights = Array("l143|100|00946a","l44|100|002dd1","l11|100|007985","l12|100|006a94","l13|100|005ba3","l14|100|004ab3","l15|100|003bc3","l05|100|009569","l07|100|008678","l08|100|007688","l09|100|006797","l10|100|0058a6","l06|100|00738b","l04|100|00827c","l03|100|00926c","l02|100|00a15d","l01|100|00b34b","l60|100|00ce30","l49|100|00708e","l59|100|002ed0","l65|100|00e816","l32|100|00ee10","l31|100|00f707","l30|100|00d32b","l41|100|00c539","l40|100|00bc42","l39|100|00b44a","l38|100|00ab53","l37|100|00a15d","l25|100|00ea14","l24|100|000000","l45|100|0042bc","l48|100|0011ed","l84|100|0022db","l83|100|000af4","l95|100|00bb43","l145|100|0022db","l133|100|000bf3","l128|100|00ef0f","l126|100|00bf3f","l110|100|00728b","l109|100|0058a6","l118|100|00936b","l117|100|00c836","l93|100|00bd41","l98|100|00906e")
	End With
	With .AddStep(2.33333, Null, Null)
		.Lights = Array("l143|100|00718d","l44|100|0046b8","l11|100|008e70","l12|100|007c82","l13|100|006c92","l14|100|005aa4","l15|100|0047b7","l05|100|00a856","l07|100|009668","l08|100|00847a","l09|100|00738b","l10|100|00619d","l06|100|007a84","l04|100|008b73","l03|100|009d61","l02|100|00af4f","l01|100|00c33b","l60|100|00bc42","l49|100|00609e","l59|100|001fdf","l65|100|00f10d","l32|100|00ea14","l31|100|00eb13","l30|100|00b945","l41|100|00c03e","l40|100|00b34b","l39|100|00a757","l38|100|009a64","l37|100|008e70","l26|100|00f806","l25|100|00cc32","l45|100|00619d","l21|100|0006f8","l48|100|0000ff","l46|100|00e816","l84|100|0032cb","l83|100|0014ea","l95|100|00a35b","l145|100|0037c7","l133|100|0032cc","l128|100|00d826","l126|100|009965","l110|100|0051ad","l109|100|0042bc","l105|100|0009f4","l118|100|00ab53","l117|100|00db23","l93|100|00c23c","l98|100|008579")
	End With
	With .AddStep(2.36667, Null, Null)
		.Lights = Array("l143|100|004cb2","l44|100|00639b","l11|100|00a45a","l12|100|00906e","l13|100|007d81","l14|100|006995","l15|100|0055a9","l05|100|00bb43","l07|100|00a757","l08|100|00926c","l09|100|007f7f","l10|100|006b93","l06|100|00817d","l04|100|00946a","l03|100|00a955","l02|100|00bd41","l01|100|00d22c","l60|100|00a856","l49|100|004eb0","l59|100|0010ee","l65|100|00f806","l32|100|00e519","l31|100|00dd21","l30|100|009b63","l41|100|00b945","l40|100|00a955","l39|100|009965","l38|100|008975","l37|100|007985","l26|100|00e519","l25|100|00ac52","l45|100|00817d","l21|100|001be3","l48|100|000000","l46|100|00c33b","l47|100|00f30b","l84|100|0044b9","l83|100|0021dd","l82|100|0000ff","l95|100|008a74","l145|100|004eb0","l133|100|005da1","l128|100|00be40","l126|100|00738b","l110|100|002fcf","l109|100|002dd1","l105|100|0026d8","l118|100|00c33b","l117|100|00ed11","l93|100|00c539","l98|100|007886")
	End With
	With .AddStep(2.40000, Null, Null)
		.Lights = Array("l143|100|0027d7","l44|100|00807e","l11|100|00bb43","l12|100|00a459","l13|100|008f6e","l14|100|007a84","l15|100|00649a","l05|100|00cd31","l07|100|00b845","l08|100|00a15d","l09|100|008c72","l10|100|007688","l06|100|008876","l04|100|009d61","l03|100|00b44a","l02|100|00ca34","l01|100|00e11d","l60|100|00936b","l49|100|003dc1","l59|100|0003fc","l65|100|00fe01","l32|100|00dd21","l31|100|00cc32","l30|100|007d81","l41|100|00b24c","l40|100|009d61","l39|100|008a74","l38|100|007787","l37|100|00649a","l27|100|00f806","l26|100|00d12d","l25|100|008975","l45|100|00a25b","l21|100|0032cc","l22|100|0000ff","l20|100|0008f6","l46|100|009a64","l47|100|00d925","l84|100|0059a5","l83|100|002fcf","l82|100|0008f6","l95|100|00708e","l137|100|00e01e","l145|100|006797","l133|100|008975","l128|100|00a15d","l126|100|004ab4","l110|100|000df1","l109|100|0018e6","l105|100|0045b9","l104|100|001de1","l102|100|0014ea","l118|100|00db23","l117|100|00ff00","l93|100|00c836","l98|100|006b93")
	End With
	With .AddStep(2.43333, Null, Null)
		.Lights = Array("l143|100|0002fd","l44|100|009e60","l11|100|00d12d","l12|100|00ba44","l13|100|00a25c","l14|100|008b73","l15|100|00748a","l05|100|00df1f","l07|100|00c835","l08|100|00b14d","l09|100|009965","l10|100|00827c","l06|100|00906e","l04|100|00a658","l03|100|00be40","l02|100|00d628","l01|100|00ef0f","l60|100|007d81","l49|100|002dd1","l59|100|000000","l65|100|000000","l32|100|00d32b","l31|100|00ba44","l30|100|005da1","l41|100|00a855","l40|100|00916c","l39|100|007b83","l38|100|006599","l37|100|004db1","l27|100|00ed11","l26|100|00ba44","l25|100|006599","l45|100|00c539","l21|100|004bb3","l22|100|001ae3","l20|100|001de1","l46|100|00718d","l47|100|00bc42","l84|100|006e90","l83|100|0040be","l82|100|0015e9","l95|100|0055a9","l137|100|00bd41","l145|100|00817d","l133|100|00b747","l128|100|00837b","l126|100|0022dc","l110|100|000000","l109|100|0005f9","l105|100|006697","l104|100|0042bc","l102|100|0042bc","l101|100|0017e7","l118|100|00f20c","l117|100|000000","l93|100|00ca34","l98|100|005ea0")
	End With
	With .AddStep(2.46667, Null, Null)
		.Lights = Array("l143|100|000000","l44|100|00bd41","l11|100|00e618","l12|100|00cd30","l13|100|00b648","l14|100|009c62","l15|100|00847a","l05|100|00ef0f","l07|100|00d826","l08|100|00bf3f","l09|100|00a658","l10|100|008d70","l06|100|009767","l04|100|00af4f","l03|100|00c836","l02|100|00e11d","l01|100|00fb03","l34|100|00e716","l60|100|006797","l19|100|00de20","l18|100|00ed11","l49|100|001ee0","l32|100|00c836","l31|100|00a559","l30|100|003dc1","l41|100|009e60","l40|100|008579","l39|100|006b93","l38|100|0052ac","l37|100|0038c6","l27|100|00df1f","l26|100|00a05d","l25|100|0041bd","l45|100|00e618","l21|100|006797","l22|100|0039c5","l20|100|0034ca","l58|100|0001fd","l55|100|0000ff","l46|100|0046b8","l47|100|009c62","l84|100|00847a","l83|100|0052ac","l82|100|0023db","l95|100|0039c5","l137|100|009668","l145|100|009b63","l133|100|00e519","l128|100|006599","l126|100|000000","l109|100|000000","l105|100|008975","l104|100|006a94","l102|100|00748a","l101|100|000000","l118|100|000000","l98|100|0051ad")
	End With
	With .AddStep(2.50000, Null, Null)
		.Lights = Array("l44|100|00da24","l11|100|00f905","l12|100|00e01e","l13|100|00c737","l14|100|00ae50","l15|100|00946a","l05|100|00fe01","l07|100|00e519","l08|100|00cc32","l09|100|00b34b","l10|100|009965","l06|100|009e60","l04|100|00b747","l03|100|00d02e","l02|100|00ea14","l01|100|000000","l53|100|0017e7","l62|100|00e21c","l34|100|00ca34","l60|100|004faf","l19|100|00b24c","l18|100|00bc42","l49|100|0010ed","l32|100|00bb43","l31|100|00906e","l30|100|001ee0","l41|100|00946a","l40|100|007886","l39|100|005ca2","l38|100|0040be","l37|100|0024da","l27|100|00ce30","l26|100|008777","l25|100|001de1","l45|100|000000","l51|100|0001fd","l21|100|00837b","l22|100|0059a4","l20|100|004db1","l58|100|002ad4","l57|100|0001fe","l55|100|0021dd","l46|100|001ce2","l47|100|007d81","l64|100|00ec12","l84|100|009a64","l83|100|006698","l82|100|0034ca","l81|100|0000ff","l95|100|0020de","l137|100|006f8f","l145|100|00b648","l133|100|000000","l128|100|0046b8","l105|100|00ac52","l104|100|00926c","l102|100|00a559","l93|100|00c935","l98|100|0045b9")
	End With
	With .AddStep(2.53333, Null, Null)
		.Lights = Array("l44|100|00f608","l11|100|000000","l12|100|00f10d","l13|100|00d826","l14|100|00bf3f","l15|100|00a45a","l05|100|000000","l07|100|00f10d","l08|100|00d826","l09|100|00bf3f","l10|100|00a45a","l06|100|00a45a","l04|100|00be40","l03|100|00d727","l02|100|00f10d","l53|100|0036c7","l35|100|00ef0f","l62|100|00be40","l34|100|00ab53","l60|100|003ac4","l19|100|00837b","l18|100|000000","l49|100|0005f9","l65|100|00fd01","l32|100|00ad51","l31|100|007a84","l30|100|0001fe","l41|100|008975","l40|100|006b93","l39|100|004cb2","l38|100|002fcf","l37|100|0012ec","l27|100|00bd41","l26|100|006d91","l25|100|000000","l51|100|0019e5","l21|100|009f5f","l22|100|007a83","l20|100|006896","l58|100|0055a9","l57|100|002ad4","l56|100|0001fd","l55|100|0045b9","l54|100|001be2","l46|100|000000","l47|100|005da1","l64|100|00d529","l84|100|00b04d","l83|100|007a84","l82|100|0046b8","l81|100|000fef","l95|100|0008f6","l137|100|0047b7","l145|100|00cf2f","l128|100|0028d6","l105|100|00ce30","l104|100|00ba44","l102|100|00d727","l93|100|00c737","l98|100|003ac4")
	End With
	With .AddStep(2.56667, Null, Null)
		.Lights = Array("l44|100|000000","l12|100|00ff00","l13|100|00e717","l14|100|00ce30","l15|100|00b549","l07|100|00fb03","l08|100|00e21c","l09|100|00c935","l10|100|00b04e","l06|100|00ab53","l04|100|00c33b","l03|100|00dc22","l02|100|00f608","l53|100|0058a6","l35|100|00d02e","l62|100|009767","l34|100|008975","l92|100|00f806","l60|100|0025d8","l19|100|0054aa","l49|100|000000","l65|100|00f608","l32|100|009d61","l31|100|00649a","l30|100|000000","l41|100|007d81","l40|100|005f9f","l39|100|003ec0","l38|100|0020de","l37|100|0001fe","l27|100|00a955","l26|100|0053ab","l51|100|0033cb","l52|100|0018e6","l21|100|00bc42","l22|100|009c62","l20|100|00837b","l58|100|00817d","l57|100|0056a8","l56|100|002cd2","l55|100|006c91","l54|100|0041bd","l47|100|003cc2","l64|100|00bd41","l84|100|00c539","l83|100|008e70","l82|100|005aa4","l81|100|0022dc","l95|100|000000","l137|100|0021dd","l145|100|00e717","l128|100|000cf2","l105|100|00ef0f","l104|100|00e11d","l102|100|000000","l101|100|00e21c","l93|100|00c33b","l98|100|0030ce")
	End With
	With .AddStep(2.60000, Null, Null)
		.Lights = Array("l70|100|00f509","l69|100|00ea14","l12|100|000000","l13|100|00f40a","l14|100|00db23","l15|100|00c33b","l07|100|000000","l08|100|00ea14","l09|100|00d22c","l10|100|00ba44","l06|100|00b04e","l04|100|00c737","l03|100|00df1f","l02|100|00f805","l53|100|007a84","l35|100|00af4f","l62|100|00718d","l34|100|006895","l92|100|00e01e","l60|100|0013eb","l19|100|0026d8","l18|100|0022dc","l65|100|00ed11","l32|100|008d71","l31|100|004eb0","l41|100|00728c","l40|100|0052ac","l39|100|0031cd","l38|100|0012ec","l37|100|000000","l27|100|009569","l26|100|003ac4","l50|100|0017e7","l51|100|004eb0","l52|100|0039c5","l21|100|00d628","l22|100|00bd40","l20|100|009e60","l58|100|00ad51","l57|100|00827c","l56|100|0059a5","l55|100|00936b","l54|100|006995","l47|100|001ee0","l64|100|00a25c","l84|100|00d925","l83|100|00a25c","l82|100|006e90","l81|100|0036c8","l80|100|0003fb","l137|100|000000","l145|100|00fe01","l128|100|000000","l105|100|000000","l104|100|000000","l101|100|000000","l93|100|00bd41","l98|100|0028d6")
	End With
	With .AddStep(2.63333, Null, Null)
		.Lights = Array("l70|100|00c737","l69|100|00c03e","l13|100|00ff00","l14|100|00e717","l15|100|00d02e","l08|100|00f00e","l09|100|00d925","l10|100|00c23c","l06|100|00b44a","l04|100|00ca34","l03|100|00e01e","l02|100|00f905","l63|100|00fa04","l53|100|009c62","l35|100|008d71","l62|100|004bb3","l34|100|0047b7","l92|100|00c737","l60|100|0003fc","l19|100|000000","l18|100|000000","l65|100|00e21c","l33|100|00ef0e","l32|100|007d81","l31|100|003ac4","l29|100|0000ff","l41|100|006797","l40|100|0047b7","l39|100|0026d8","l38|100|0007f7","l28|100|00f20c","l27|100|00807e","l26|100|0022dc","l50|100|0035c9","l51|100|006b92","l52|100|005da1","l21|100|00ef0f","l22|100|00dd21","l20|100|00b945","l58|100|00d826","l57|100|00ae50","l56|100|008579","l55|100|00ba44","l54|100|00906e","l47|100|0002fd","l64|100|008777","l84|100|00eb13","l83|100|00b648","l82|100|00837b","l81|100|004cb2","l80|100|001ae4","l145|100|000000","l107|100|0007f7","l106|100|0023db","l93|100|00b747","l98|100|0022dc")
	End With
	With .AddStep(2.66667, Null, Null)
		.Lights = Array("l70|100|009767","l69|100|00946a","l13|100|000000","l14|100|00f10d","l15|100|00dc22","l08|100|00f40a","l09|100|00de20","l10|100|00c935","l06|100|00b747","l04|100|00cb33","l02|100|00f607","l63|100|00da24","l53|100|00bd41","l35|100|006b93","l62|100|0026d8","l34|100|0029d5","l92|100|00ac52","l60|100|000000","l65|100|00d628","l33|100|00da24","l32|100|006d90","l31|100|0028d6","l29|100|0014ea","l41|100|005da1","l40|100|003dc1","l39|100|001de1","l38|100|0000ff","l28|100|00d825","l27|100|006c92","l26|100|000df1","l24|100|0000ff","l50|100|0054aa","l51|100|008876","l52|100|00807e","l21|100|000000","l22|100|00fa04","l20|100|00d12c","l58|100|000000","l57|100|00d925","l56|100|00b24b","l55|100|00de20","l54|100|00b747","l47|100|000000","l64|100|006d91","l84|100|00fa04","l83|100|00c836","l82|100|009767","l81|100|00639b","l80|100|0033cb","l136|100|001ae4","l107|100|002cd1","l106|100|004db1","l93|100|00af4f","l98|100|001de1")
	End With
	With .AddStep(2.70000, Null, Null)
		.Lights = Array("l142|100|00f608","l70|100|006896","l69|100|006995","l14|100|00f806","l15|100|00e519","l08|100|00f509","l09|100|00e21c","l10|100|00cf2f","l06|100|00b846","l04|100|00ca34","l03|100|00dd20","l02|100|00f20c","l63|100|00b846","l53|100|00dc22","l35|100|0049b5","l62|100|0004fa","l34|100|000cf2","l92|100|00906e","l65|100|00c836","l33|100|00c33b","l32|100|005ea0","l31|100|0017e7","l29|100|002ad4","l41|100|0053ab","l40|100|0035c9","l39|100|0016e8","l38|100|000000","l28|100|00bd41","l27|100|0059a5","l26|100|000000","l24|100|0012ec","l43|100|001ee0","l50|100|00748a","l51|100|00a35a","l52|100|00a35b","l22|100|000000","l20|100|00e816","l57|100|000000","l56|100|00dd21","l55|100|000000","l54|100|00dc22","l64|100|0052ac","l84|100|000000","l83|100|00d826","l82|100|00ab53","l81|100|007a84","l80|100|004db1","l136|100|0038c6","l107|100|0053ab","l106|100|007985","l93|100|00a658","l98|100|001ae3")
	End With
	With .AddStep(2.73333, Null, Null)
		.Lights = Array("l142|100|00d727","l70|100|0039c4","l69|100|003ec0","l14|100|00fe01","l15|100|00ed11","l09|100|00e31b","l10|100|00d32b","l06|100|00b945","l04|100|00c935","l03|100|00d925","l02|100|00ec12","l01|100|00fd01","l63|100|009569","l53|100|00f905","l35|100|002ad3","l62|100|000000","l34|100|000000","l92|100|007589","l59|100|000af4","l65|100|00b846","l33|100|00ab53","l32|100|004faf","l31|100|0009f5","l29|100|0042bc","l41|100|004ab4","l40|100|002ed0","l39|100|0010ee","l28|100|00a05e","l27|100|0045b9","l24|100|0026d8","l43|100|0040bd","l50|100|00926c","l51|100|00bf3f","l52|100|00c538","l20|100|00fd01","l56|100|000000","l54|100|00ff00","l48|100|0001fe","l64|100|003ac4","l42|100|001ce2","l83|100|00e618","l82|100|00bd41","l81|100|00916d","l80|100|006896","l136|100|0058a6","l120|100|00df1f","l107|100|007a84","l106|100|00a35b","l93|100|009d61","l98|100|001ae4")
	End With
	With .AddStep(2.76667, Null, Null)
		.Lights = Array("l142|100|00b648","l71|100|0005f9","l70|100|000ef0","l69|100|0017e7","l14|100|00ff00","l15|100|00f20c","l07|100|00ff00","l08|100|00f20c","l10|100|00d529","l04|100|00c638","l03|100|00d42a","l02|100|00e41a","l01|100|00f20c","l63|100|00738b","l53|100|000000","l35|100|000ef0","l92|100|005ba3","l59|100|0018e6","l65|100|00a856","l33|100|00926c","l32|100|0042bc","l31|100|000000","l29|100|005aa4","l41|100|0042bc","l40|100|0028d5","l39|100|000df1","l28|100|008479","l27|100|0034ca","l24|100|003cc2","l43|100|00649a","l50|100|00b14d","l51|100|00d826","l52|100|00e519","l20|100|000000","l54|100|000000","l48|100|0014ea","l64|100|0023db","l42|100|0043bb","l83|100|00f20c","l82|100|00ce30","l81|100|00a658","l80|100|00827c","l136|100|007688","l120|100|00bc42","l119|100|00f508","l108|100|0009f5","l107|100|009f5f","l106|100|00cc32","l93|100|00946a","l98|100|001be3")
	End With
	With .AddStep(2.80000, Null, Null)
		.Lights = Array("l142|100|009569","l71|100|000000","l70|100|000000","l69|100|000000","l14|100|000000","l15|100|00f509","l07|100|00f806","l08|100|00ee10","l09|100|00e21c","l10|100|00d628","l06|100|00b746","l04|100|00c23c","l03|100|00cd31","l02|100|00da24","l01|100|00e519","l63|100|0053ab","l35|100|000000","l92|100|0041bd","l49|100|0005f9","l59|100|0028d6","l65|100|009866","l33|100|007b83","l32|100|0036c8","l29|100|00728c","l41|100|003cc2","l40|100|0025d9","l39|100|000cf2","l28|100|006a94","l27|100|0025d9","l24|100|0052ac","l43|100|008678","l50|100|00cd31","l51|100|00ee10","l52|100|000000","l48|100|0027d7","l64|100|000fef","l42|100|006a93","l83|100|00fd02","l82|100|00dd21","l81|100|00bb43","l80|100|009b63","l136|100|00946a","l120|100|009965","l119|100|00db23","l108|100|002ad4","l107|100|00c33b","l106|100|00f20c","l117|100|00f30b","l93|100|008b73","l98|100|001de1")
	End With
	With .AddStep(2.83333, Null, Null)
		.Lights = Array("l142|100|007688","l14|100|00ff00","l15|100|00f707","l05|100|00f707","l07|100|00ef0f","l08|100|00e816","l09|100|00de20","l06|100|00b549","l04|100|00bd41","l03|100|00c638","l02|100|00d02e","l01|100|00d826","l63|100|0034ca","l92|100|002bd3","l49|100|000fef","l59|100|0038c6","l65|100|008876","l33|100|00649a","l32|100|002cd2","l29|100|008876","l41|100|0037c7","l40|100|0022dc","l28|100|004faf","l27|100|0018e6","l24|100|006895","l43|100|00a757","l50|100|00e618","l51|100|000000","l48|100|003cc2","l64|100|000000","l42|100|008f6e","l83|100|000000","l82|100|00ea14","l81|100|00ce30","l80|100|00b34b","l136|100|00b14d","l120|100|007886","l119|100|00c03d","l108|100|004bb3","l107|100|00e41a","l106|100|000000","l117|100|00e21c","l93|100|00827c","l98|100|0021dd")
	End With
	With .AddStep(2.86667, Null, Null)
		.Lights = Array("l142|100|0059a5","l14|100|00fc02","l15|100|00f608","l05|100|00eb13","l07|100|00e518","l08|100|00e11d","l09|100|00da24","l10|100|00d429","l06|100|00b34b","l04|100|00b846","l03|100|00be40","l02|100|00c539","l01|100|00ca34","l63|100|0019e5","l92|100|0017e7","l49|100|001ae4","l59|100|0049b5","l65|100|007985","l33|100|004faf","l32|100|0024da","l29|100|009e60","l41|100|0033cb","l39|100|000ef0","l28|100|0038c6","l27|100|000df1","l24|100|007d81","l43|100|00c638","l50|100|00fd01","l48|100|0050ae","l42|100|00b34b","l82|100|00f40a","l81|100|00de20","l80|100|00c836","l136|100|00cb33","l120|100|0058a6","l119|100|00a658","l109|100|000bf3","l108|100|006c92","l107|100|000000","l117|100|00d02e","l93|100|007984","l98|100|0026d8")
	End With
	With .AddStep(2.90000, Null, Null)
		.Lights = Array("l142|100|003ec0","l12|100|00ff00","l13|100|00fb03","l14|100|00f806","l15|100|00f509","l05|100|00de20","l07|100|00db23","l08|100|00d925","l09|100|00d529","l10|100|00d22c","l06|100|00b04e","l04|100|00b24c","l03|100|00b548","l02|100|00ba44","l01|100|00bc42","l63|100|0000fe","l92|100|0005f9","l49|100|0025d9","l59|100|0059a5","l65|100|006c92","l33|100|003cc2","l32|100|001de1","l29|100|00b24c","l41|100|0031cd","l39|100|0011ed","l38|100|0002fc","l28|100|0023db","l27|100|0004fa","l24|100|00906e","l43|100|00e11d","l50|100|000000","l48|100|00639a","l42|100|00d32b","l82|100|00fd01","l81|100|00ec12","l80|100|00db23","l136|100|00e21c","l120|100|003bc3","l119|100|008d71","l109|100|001ae4","l108|100|008975","l118|100|00ff00","l117|100|00bf3f","l93|100|00728c","l98|100|002cd2")
	End With
	With .AddStep(2.93333, Null, Null)
		.Lights = Array("l142|100|0027d7","l11|100|00f608","l12|100|00f509","l13|100|00f30b","l14|100|00f30b","l15|100|00f20c","l05|100|00d22c","l07|100|00d12d","l08|100|00d22c","l09|100|00d02e","l10|100|00cf2f","l06|100|00ad51","l04|100|00ad51","l03|100|00ae50","l02|100|00b04e","l01|100|00af4f","l63|100|000000","l92|100|000000","l49|100|0030ce","l59|100|006896","l65|100|00609e","l33|100|002cd2","l32|100|0018e6","l29|100|00c33b","l41|100|002fcf","l40|100|0023db","l39|100|0015e9","l38|100|0009f5","l28|100|0011ed","l27|100|000000","l24|100|00a15d","l43|100|00f806","l48|100|007589","l42|100|00ee10","l82|100|000000","l81|100|00f707","l80|100|00eb13","l136|100|00f509","l120|100|0022dc","l119|100|007787","l109|100|002ad4","l108|100|00a45a","l118|100|00ef0e","l117|100|00af4f","l93|100|006b93","l98|100|0032cc")
	End With
	With .AddStep(2.96667, Null, Null)
		.Lights = Array("l142|100|0013eb","l11|100|00ea14","l12|100|00ec12","l13|100|00ec12","l14|100|00ee10","l15|100|00ef0f","l05|100|00c638","l07|100|00c836","l08|100|00cb33","l09|100|00cb33","l10|100|00cc32","l06|100|00a955","l04|100|00a757","l03|100|00a658","l02|100|00a559","l01|100|00a35b","l49|100|003ac4","l59|100|007589","l65|100|0055a9","l33|100|001fdf","l32|100|0014ea","l29|100|00d12d","l41|100|002ed0","l40|100|0025d9","l39|100|0019e5","l38|100|000fef","l37|100|0006f8","l28|100|0003fb","l24|100|00b14d","l43|100|000000","l48|100|00847a","l42|100|000000","l81|100|000000","l80|100|00f905","l136|100|000000","l120|100|000df1","l119|100|00649a","l109|100|0037c7","l108|100|00bc42","l118|100|00e11c","l117|100|00a05e","l93|100|006599","l98|100|0037c7")
	End With
	With .AddStep(3.00000, Null, Null)
		.Lights = Array("l142|100|0004fa","l11|100|00e11d","l12|100|00e41a","l13|100|00e618","l14|100|00e915","l15|100|00ec12","l05|100|00bd41","l07|100|00c03e","l08|100|00c43a","l09|100|00c638","l10|100|00c935","l06|100|00a658","l04|100|00a35b","l03|100|009f5e","l02|100|009e60","l01|100|009964","l49|100|0042bc","l59|100|007f7f","l65|100|004cb2","l33|100|0014ea","l32|100|0011ec","l29|100|00dc22","l40|100|0026d8","l39|100|001de1","l38|100|0015e9","l37|100|000df1","l28|100|000000","l24|100|00bc41","l48|100|00906e","l80|100|000000","l120|100|000000","l119|100|0054aa","l109|100|0043bb","l108|100|00ce30","l118|100|00d628","l117|100|00946a","l93|100|00609e","l98|100|003cc2")
	End With
	With .AddStep(3.03333, Null, Null)
		.Lights = Array("l142|100|000000","l11|100|00da24","l12|100|00de20","l13|100|00e11d","l14|100|00e519","l15|100|00ea14","l05|100|00b648","l07|100|00ba43","l08|100|00c03e","l09|100|00c33b","l10|100|00c737","l06|100|00a45a","l04|100|00a05e","l03|100|009b63","l02|100|009866","l01|100|00936b","l49|100|0048b6","l59|100|008678","l65|100|0047b7","l33|100|000ef0","l32|100|0010ee","l29|100|00e31b","l40|100|0028d6","l39|100|0020de","l38|100|001ae4","l37|100|0013eb","l24|100|00c43a","l48|100|009865","l119|100|0049b5","l110|100|0004fa","l109|100|004bb3","l108|100|00da24","l118|100|00cd31","l117|100|008c72","l93|100|005da1","l98|100|0040be")
	End With
	With .AddStep(3.06667, Null, Null)
		.Lights = Array("l11|100|00d727","l12|100|00dc22","l13|100|00df1f","l14|100|00e41a","l15|100|00e915","l05|100|00b34b","l07|100|00b846","l08|100|00be40","l09|100|00c23c","l10|100|00c638","l06|100|00a35b","l04|100|009e60","l03|100|009a64","l02|100|009668","l01|100|00906e","l49|100|004bb3","l59|100|008975","l65|100|0045b9","l33|100|000bf3","l29|100|00e618","l39|100|0021dd","l38|100|001be3","l37|100|0015e9","l24|100|00c737","l48|100|009c62","l119|100|0045b9","l110|100|0008f6","l109|100|004eb0","l108|100|00df1f","l118|100|00ca34","l117|100|008975","l93|100|005ca2","l98|100|0041bd")
	End With
	With .AddStep(3.10000, Null, Null)
		.Lights = Array("l11|100|00d42a","l12|100|00d826","l13|100|00dc22","l14|100|00e01d","l15|100|00e519","l05|100|00b04e","l07|100|00b549","l08|100|00bb43","l09|100|00be40","l10|100|00c33b","l06|100|00a05e","l04|100|009b63","l03|100|009668","l02|100|00936b","l01|100|008d71","l49|100|0047b7","l59|100|008678","l65|100|0041bd","l33|100|0008f6","l32|100|000cf2","l29|100|00e31b","l41|100|002bd3","l40|100|0025d9","l39|100|001ee0","l38|100|0018e6","l37|100|0012ec","l24|100|00c43a","l48|100|009965","l119|100|0041bd","l110|100|0005f9","l109|100|004bb3","l108|100|00dc22","l118|100|00c737","l117|100|008678","l93|100|0059a5","l98|100|003ec0")
	End With
	With .AddStep(3.13333, Null, Null)
		.Lights = Array("l11|100|00ca34","l12|100|00cf2f","l13|100|00d22c","l14|100|00d727","l15|100|00dc22","l05|100|00a658","l07|100|00ab52","l08|100|00b14d","l09|100|00b549","l10|100|00ba44","l06|100|009668","l04|100|00926c","l03|100|008d71","l02|100|008975","l01|100|00837b","l49|100|003ec0","l59|100|007c82","l65|100|0038c6","l33|100|0000ff","l32|100|0003fb","l29|100|00d924","l41|100|0022dc","l40|100|001ce2","l39|100|0015e9","l38|100|000fef","l37|100|0009f5","l24|100|00bb43","l48|100|008f6f","l80|100|00ff00","l119|100|0038c6","l110|100|000000","l109|100|0041bc","l108|100|00d22c","l118|100|00bd41","l117|100|007c82","l93|100|004eaf","l98|100|0034c9")
	End With
	With .AddStep(3.16667, Null, Null)
		.Lights = Array("l11|100|00bc42","l12|100|00c13d","l13|100|00c43a","l14|100|00c935","l15|100|00ce30","l05|100|009866","l07|100|009c61","l08|100|00a25c","l09|100|00a658","l10|100|00ac52","l06|100|008876","l04|100|00847a","l03|100|007f7f","l02|100|007b83","l01|100|007588","l49|100|0030ce","l59|100|006e90","l65|100|002ad4","l33|100|000000","l32|100|000000","l29|100|00cc32","l41|100|0014ea","l40|100|000ef0","l39|100|0007f7","l38|100|0001fe","l37|100|000000","l24|100|00ad51","l48|100|00817d","l84|100|00f20b","l83|100|00f20b","l82|100|00f20b","l81|100|00f20b","l80|100|00f10d","l119|100|002ad4","l109|100|0033ca","l108|100|00c43a","l118|100|00af4f","l117|100|006e90","l93|100|0040bd","l98|100|0026d7")
	End With
	With .AddStep(3.20000, Null, Null)
		.Lights = Array("l44|100|00fd02","l11|100|00ab53","l12|100|00af4f","l13|100|00b34b","l14|100|00b846","l15|100|00bc42","l05|100|008678","l07|100|008b73","l08|100|00916d","l09|100|00946a","l10|100|009965","l06|100|007787","l04|100|00728c","l03|100|006d91","l02|100|006a94","l01|100|00649a","l49|100|001edf","l59|100|005da1","l65|100|0019e5","l29|100|00ba44","l41|100|0002fc","l40|100|000000","l39|100|000000","l38|100|000000","l24|100|009a64","l45|100|00fd02","l43|100|00fd02","l48|100|00708e","l42|100|00fd02","l84|100|00e11d","l83|100|00e11d","l82|100|00e11d","l81|100|00e11d","l80|100|00e01e","l136|100|00f10c","l145|100|00f30b","l119|100|0019e5","l109|100|0022dc","l108|100|00b34b","l118|100|009d61","l117|100|005da1","l93|100|002fcf","l98|100|0015e9")
	End With
	With .AddStep(3.23333, Null, Null)
		.Lights = Array("l44|100|00e816","l11|100|009569","l12|100|009a64","l13|100|009e60","l14|100|00a25c","l15|100|00a856","l05|100|00728c","l07|100|007787","l08|100|007d81","l09|100|00807e","l10|100|008579","l06|100|00639b","l04|100|005ea0","l03|100|0059a5","l02|100|0056a8","l01|100|004faf","l49|100|000af4","l59|100|0048b6","l65|100|0004fa","l29|100|00a559","l41|100|000000","l24|100|008678","l45|100|00e816","l43|100|00e816","l50|100|00ff00","l51|100|00ff00","l21|100|00fe01","l20|100|00fe01","l48|100|005ba2","l42|100|00e816","l84|100|00cd31","l83|100|00cd31","l82|100|00cd31","l81|100|00cd31","l80|100|00cc32","l136|100|00dd21","l145|100|00de20","l119|100|0004fa","l109|100|000ef0","l108|100|009e60","l118|100|008876","l117|100|0048b6","l93|100|001be3","l98|100|0001fe")
	End With
	With .AddStep(3.26667, Null, Null)
		.Lights = Array("l44|100|00d12d","l11|100|007f7f","l12|100|00847a","l13|100|008777","l14|100|008c72","l15|100|00906d","l05|100|005ba3","l07|100|00609e","l08|100|006698","l09|100|006a94","l10|100|006e90","l06|100|004bb3","l04|100|0046b8","l03|100|0042bc","l02|100|003ec0","l01|100|0038c6","l49|100|000000","l59|100|0031cd","l65|100|000000","l29|100|008e70","l24|100|006f8f","l45|100|00d12d","l43|100|00d12d","l50|100|00e816","l51|100|00e816","l21|100|00e717","l20|100|00e717","l48|100|0044ba","l42|100|00d12d","l84|100|00b648","l83|100|00b648","l82|100|00b648","l81|100|00b648","l80|100|00b549","l136|100|00c737","l145|100|00c836","l119|100|000000","l109|100|000000","l108|100|008777","l118|100|00728c","l117|100|0031cd","l93|100|0004fa","l98|100|000000")
	End With
	With .AddStep(3.30000, Null, Null)
		.Lights = Array("l44|100|00b945","l11|100|006698","l12|100|006b93","l13|100|006f8f","l14|100|00738b","l15|100|007886","l05|100|0042bc","l07|100|0047b7","l08|100|004db1","l09|100|0050ae","l10|100|0056a8","l06|100|0033cb","l04|100|002ed0","l03|100|0029d5","l02|100|0026d8","l01|100|0020de","l53|100|00fe00","l59|100|0019e5","l29|100|007688","l24|100|0057a7","l45|100|00b945","l43|100|00b945","l50|100|00cf2f","l51|100|00cf2f","l52|100|00fb03","l21|100|00ce30","l22|100|00fa04","l20|100|00ce30","l48|100|002cd2","l42|100|00b945","l84|100|009d61","l83|100|009d61","l82|100|009d61","l81|100|009d61","l80|100|009c62","l136|100|00ae50","l145|100|00af4e","l108|100|006f8f","l107|100|00ed11","l105|100|00ed11","l118|100|0059a5","l117|100|0019e5","l93|100|000000")
	End With
	With .AddStep(3.33333, Null, Null)
		.Lights = Array("l44|100|009e60","l11|100|004cb2","l12|100|0050ae","l13|100|0055a9","l14|100|005aa4","l15|100|005ea0","l05|100|0028d6","l07|100|002dd1","l08|100|0033cb","l09|100|0036c8","l10|100|003bc3","l06|100|0019e5","l04|100|0014ea","l03|100|000fef","l02|100|000cf2","l01|100|0006f8","l53|100|00e41a","l59|100|0000ff","l29|100|005ca2","l24|100|003cc2","l45|100|009e60","l43|100|009e60","l50|100|00b548","l51|100|00b548","l52|100|00e11d","l21|100|00b44a","l22|100|00e01e","l20|100|00b44a","l48|100|0012ec","l42|100|009e60","l84|100|00837b","l83|100|00837b","l82|100|00837b","l81|100|00837b","l80|100|00827c","l136|100|00936b","l145|100|009569","l133|100|00f10d","l108|100|0055a9","l107|100|00d32b","l105|100|00d32b","l118|100|003fbf","l117|100|0000ff")
	End With
	With .AddStep(3.36667, Null, Null)
		.Lights = Array("l44|100|00837b","l11|100|0030ce","l12|100|0035c9","l13|100|0039c5","l14|100|003dc0","l15|100|0042bc","l05|100|000df1","l07|100|0012ec","l08|100|0018e6","l09|100|001be3","l10|100|0020de","l06|100|000000","l04|100|000000","l03|100|000000","l02|100|000000","l01|100|000000","l53|100|00c835","l59|100|000000","l29|100|0040be","l24|100|0021dd","l45|100|00837b","l43|100|00837b","l50|100|009965","l51|100|009965","l52|100|00c638","l21|100|009866","l22|100|00c539","l20|100|009866","l55|100|00ec12","l54|100|00ec12","l48|100|000000","l42|100|00837b","l84|100|006896","l83|100|006896","l82|100|006896","l81|100|006896","l80|100|006797","l136|100|007886","l145|100|007984","l133|100|00d529","l108|100|0039c5","l107|100|00b846","l106|100|00e618","l105|100|00b846","l104|100|00e618","l118|100|0023db","l117|100|000000")
	End With
	With .AddStep(3.40000, Null, Null)
		.Lights = Array("l44|100|006797","l11|100|0014ea","l12|100|0019e5","l13|100|001de1","l14|100|0021dd","l15|100|0026d8","l05|100|000000","l07|100|000000","l08|100|000000","l09|100|0000ff","l10|100|0004fa","l53|100|00ac52","l29|100|0024da","l24|100|0005f9","l45|100|006797","l43|100|006797","l50|100|007d81","l51|100|007d81","l52|100|00aa54","l21|100|007c82","l22|100|00a955","l20|100|007c82","l58|100|00f40a","l57|100|00f40a","l56|100|00f40a","l55|100|00d02e","l54|100|00d02e","l42|100|006797","l84|100|004bb3","l83|100|004bb3","l82|100|004bb3","l81|100|004bb3","l80|100|0049b4","l136|100|005ca2","l145|100|005da1","l134|100|00fd01","l133|100|00b945","l108|100|001de1","l107|100|009b63","l106|100|00ca34","l105|100|009b63","l104|100|00ca34","l118|100|0007f7")
	End With
	With .AddStep(3.43333, Null, Null)
		.Lights = Array("l44|100|0049b5","l11|100|000000","l12|100|000000","l13|100|0000ff","l14|100|0004fa","l15|100|0009f5","l09|100|000000","l10|100|000000","l53|100|008e70","l29|100|0007f7","l24|100|000000","l45|100|0049b5","l43|100|0049b5","l50|100|00609e","l51|100|00609e","l52|100|008c72","l21|100|005f9f","l22|100|008b73","l20|100|005f9f","l58|100|00d826","l57|100|00d826","l56|100|00d826","l55|100|00b34b","l54|100|00b34b","l42|100|0049b5","l84|100|002ed0","l83|100|002ed0","l82|100|002ed0","l81|100|002ed0","l80|100|002dd1","l136|100|003ec0","l145|100|0040be","l134|100|00e01e","l133|100|009b62","l108|100|0000ff","l107|100|007e80","l106|100|00ad51","l105|100|007e80","l104|100|00ad51","l103|100|00e816","l102|100|00e717","l101|100|00ff00","l100|100|00ff00","l118|100|000000")
	End With
	With .AddStep(3.46667, Null, Null)
		.Lights = Array("l44|100|002bd2","l13|100|000000","l14|100|000000","l15|100|000000","l53|100|00718d","l29|100|000000","l45|100|002bd2","l43|100|002bd2","l50|100|0042bc","l51|100|0042bc","l52|100|006f8f","l21|100|0041bd","l22|100|006d91","l20|100|0041bd","l58|100|00ba44","l57|100|00ba44","l56|100|00ba44","l55|100|00946a","l54|100|00946a","l42|100|002bd2","l84|100|0010ee","l83|100|0010ee","l82|100|0010ee","l81|100|0010ee","l80|100|000fef","l136|100|0021dd","l145|100|0022dc","l134|100|00c23c","l133|100|007e80","l108|100|000000","l107|100|00609e","l106|100|008e70","l105|100|00609e","l104|100|008e70","l103|100|00cb33","l102|100|00c935","l101|100|00e21c","l100|100|00e11d")
	End With
	With .AddStep(3.50000, Null, Null)
		.Lights = Array("l44|100|000df0","l53|100|0052ac","l45|100|000df0","l43|100|000df0","l50|100|0024da","l51|100|0024da","l52|100|0050ae","l21|100|0023db","l22|100|004eb0","l20|100|0023db","l58|100|009b63","l57|100|009b63","l56|100|009b63","l55|100|007688","l54|100|007688","l42|100|000df0","l84|100|000000","l83|100|000000","l82|100|000000","l81|100|000000","l80|100|000000","l136|100|0003fb","l145|100|0004fa","l134|100|00a35b","l133|100|00609e","l107|100|0041bd","l106|100|00708d","l105|100|0041bd","l104|100|00708d","l103|100|00ad51","l102|100|00ab53","l101|100|00c43a","l100|100|00c33b")
	End With
	With .AddStep(3.53333, Null, Null)
		.Lights = Array("l44|100|000000","l16|100|00f707","l53|100|0034ca","l45|100|000000","l43|100|000000","l50|100|0006f8","l51|100|0006f8","l52|100|0031cd","l21|100|0004fa","l22|100|0030ce","l20|100|0004fa","l58|100|007d81","l57|100|007d81","l56|100|007d81","l55|100|0058a6","l54|100|0058a6","l42|100|000000","l136|100|000000","l145|100|000000","l134|100|008579","l133|100|0041bd","l107|100|0023db","l106|100|0051ad","l105|100|0023db","l104|100|0051ad","l103|100|008d71","l102|100|008c72","l101|100|00a559","l100|100|00a45a")
	End With
	With .AddStep(3.56667, Null, Null)
		.Lights = Array("l16|100|00d925","l53|100|0015e9","l50|100|000000","l51|100|000000","l52|100|0013eb","l21|100|000000","l22|100|0012ec","l20|100|000000","l58|100|005ea0","l57|100|005ea0","l56|100|005ea0","l55|100|0039c5","l54|100|0039c5","l134|100|006698","l133|100|0022dc","l107|100|0005f9","l106|100|0033cb","l105|100|0005f9","l104|100|0033cb","l103|100|006f8f","l102|100|006e90","l101|100|008678","l100|100|008579")
	End With
	With .AddStep(3.60000, Null, Null)
		.Lights = Array("l16|100|00ba44","l53|100|000000","l52|100|000000","l22|100|000000","l58|100|003fbf","l57|100|003fbf","l56|100|003fbf","l55|100|001ae4","l54|100|001ae4","l134|100|0047b7","l133|100|0004fa","l107|100|000000","l106|100|0014ea","l105|100|000000","l104|100|0014ea","l103|100|004faf","l102|100|004eb0","l101|100|006896","l100|100|006797")
	End With
	With .AddStep(3.63333, Null, Null)
		.Lights = Array("l16|100|009b63","l58|100|0020de","l57|100|0020de","l56|100|0020de","l55|100|000000","l54|100|000000","l134|100|0028d6","l133|100|000000","l106|100|000000","l104|100|000000","l103|100|0031cd","l102|100|0030ce","l101|100|0048b6","l100|100|0047b7")
	End With
	With .AddStep(3.66667, Null, Null)
		.Lights = Array("l16|100|007d81","l58|100|0002fd","l57|100|0002fd","l56|100|0002fd","l134|100|000af4","l103|100|0012ec","l102|100|0011ed","l101|100|002ad4","l100|100|0029d5")
	End With
	With .AddStep(3.70000, Null, Null)
		.Lights = Array("l16|100|005f9f","l58|100|000000","l57|100|000000","l56|100|000000","l134|100|000000","l103|100|000000","l102|100|000000","l101|100|000cf2","l100|100|000bf3")
	End With
	With .AddStep(3.73333, Null, Null)
		.Lights = Array("l16|100|0040be","l101|100|000000","l100|100|000000")
	End With
	With .AddStep(3.76667, Null, Null)
		.Lights = Array("l16|100|0023db")
	End With
	With .AddStep(3.80000, Null, Null)
		.Lights = Array("l16|100|0007f6")
	End With
	With .AddStep(3.83333, Null, Null)
		.Lights = Array("l16|100|000000")
	End With
End With



With GlfShotProfiles("default")
	With .States("on")
			.Show = glf_ShowFlash
	End With
	With .States("off")
			.Show = glf_ShowOff
	End With
End With

With GlfShotProfiles("flash_color")
	With .States("off")
		.Show = glf_ShowOff
	End With
	With .States("on")
			.Show = glf_ShowFlashColor
	End With	
End With


'******************************************************
'*****   GLF Pin Events                            ****
'******************************************************

Const GLF_GAME_START = "game_start"
Const GLF_GAME_STARTED = "game_started"
Const GLF_GAME_OVER = "game_ended"
Const GLF_BALL_WILL_END = "ball_will_end"
Const GLF_BALL_ENDING = "ball_ending"
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
        If Err Then MsgBox("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

	Public Sub Send(commandMessage)
		If m_connected = True Then
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
        End If
    End Sub
End Class

Sub Glf_BcpSendPlayerVar(args)
    If IsNull(bcpController) Then
        Exit Sub
    End If
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
        Dim message, parameters, parameter, eventName
        For Each message in messages
            'debug.print(message.Command)
            Select Case message.Command
                case "hello"
                    bcpController.Reset
                case "monitor_start"
                    Dim category : category = message.GetValue("category")
                    If category = "player_vars" Then
                        AddPlayerStateEventListener "score", "bcp_player_var_score", "Glf_BcpSendPlayerVar", 1000, Null
                        AddPlayerStateEventListener "current_ball", "bcp_player_var_ball", "Glf_BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    eventName = message.GetValue("event")
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************


Class GlfBallHold

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device

    Private m_enabled
    Private m_balls_to_hold
    Private m_hold_devices
    Private m_balls_held
    Private m_hold_queue
    Private m_release_all_events
    Private m_release_one_events
    Private m_release_one_if_full_events

    Private m_control_events
    Private m_running
    Private m_ticks
    Private m_ticks_remaining
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_tick_interval
    Private m_starting_tick_interval
    Private m_max_value
    Private restart_on_complete
    Private m_start_running

    Public Property Get Name() : Name = m_name : End Property

    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    
    Public Property Get BallsToHold() : BallsToHold = m_balls_to_hold : End Property
    Public Property Let BallsToHold(value) : m_balls_to_hold = value : End Property

    Public Property Get HoldDevices() : HoldDevices = m_hold_devices : End Property
    Public Property Let HoldDevices(value) : m_hold_devices = value : End Property

    Public Property Get ReleaseAllEvents(): Set ReleaseAllEvents = m_release_all_events: End Property
    Public Property Let ReleaseAllEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_all_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get ReleaseOneEvents(): Set ReleaseOneEvents = m_release_one_events: End Property
    Public Property Let ReleaseOneEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_one_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get ReleaseOneIfFullEvents(): Set ReleaseOneIfFullEvents = m_release_one_if_full_events: End Property
    Public Property Let ReleaseOneIfFullEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_one_if_full_events.Add newEvent.Name, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = "ball_hold_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_balls_to_hold = 0
        m_balls_held = 0
        m_hold_devices = Array()
        Set m_hold_queue = CreateObject("Scripting.Dictionary")
        Set m_release_all_events = CreateObject("Scripting.Dictionary")
        Set m_release_one_events = CreateObject("Scripting.Dictionary")
        Set m_release_one_if_full_events = CreateObject("Scripting.Dictionary")

        Set m_base_device = (new GlfBaseModeDevice)(mode, "ball_hold", Me)
        glf_ball_holds.Add name, Me
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
        m_enabled = True
        'Add Event Listeners
        Dim device
        For Each device in m_hold_devices
            AddPinEventListener "balldevice_" & device & "_ball_enter", m_mode & "_" & name & "_hold", "BallHoldsEventHandler", m_priority, Array("hold", me, device)
        Next
        Dim evt
        For Each evt in m_release_all_events.Keys
            AddPinEventListener m_release_all_events(evt).EventName, m_mode & "_" & name & "_release_all", "BallHoldsEventHandler", m_priority, Array("release_all", me, m_release_all_events(evt))
        Next
        For Each evt in m_release_one_events.Keys
            AddPinEventListener m_release_one_events(evt).EventName, m_mode & "_" & name & "_release_one", "BallHoldsEventHandler", m_priority, Array("release_one", me, m_release_one_events(evt))
        Next
        For Each evt in m_release_one_if_full_events.Keys
            AddPinEventListener m_release_one_if_full_events(evt).EventName, m_mode & "_" & name & "_release_one_if_full", "BallHoldsEventHandler", m_priority, Array("release_one_if_full", me, m_release_one_if_full_events(evt))
        Next
    End Sub

    Public Sub Disable()
        m_enabled = False
        Dim device
        For Each device in m_hold_devices
            RemovePinEventListener "balldevice_" & device & "_ball_enter", m_mode & "_" & name & "_hold"
        Next
        Dim evt
        For Each evt in m_release_all_events.Keys
            RemovePinEventListener m_release_all_events(evt).EventName, m_mode & "_" & name & "_release_all"
        Next
        For Each evt in m_release_one_events.Keys
            RemovePinEventListener m_release_one_events(evt).EventName, m_mode & "_" & name & "_release_one"
        Next
        For Each evt in m_release_one_if_full_events.Keys
            RemovePinEventListener m_release_one_if_full_events(evt).EventName, m_mode & "_" & name & "_release_one_if_full"
        Next
    End Sub

    Public Function IsFull()
        'Return true if hold is full
        If RemainingSpaceInHold() = 0 Then
            IsFull = True
        Else
            IsFull = False
        End If
    End Function

    Public Function RemainingSpaceInHold()
        'Return the remaining capacity of the hold.
        Dim balls
        balls = m_balls_to_hold - m_balls_held
        If balls < 0 Then
            balls = 0
        End If
        RemainingSpaceInHold = balls
    End Function

    Public Function HoldBall(device, unclaimed_balls)        
        ' Handle result of _ball_enter event of hold_devices.
        If IsFull() Then
            m_base_device.Log "Cannot hold balls. Hold is full."
            HoldBall = unclaimed_balls
            Exit Function
        End If

        If unclaimed_balls <= 0 Then
            HoldBalls = unclaimed_balls
            Exit Function
        End If

        Dim capacity : capacity = RemainingSpaceInHold()
        Dim balls_to_hold
        If unclaimed_balls > capacity Then
            balls_to_hold = capacity
        Else
            balls_to_hold = unclaimed_balls
        End If
        m_balls_held = m_balls_held + balls_to_hold
        m_base_device.Log "Held " & balls_to_hold & " balls"

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls_held", balls_to_hold
            .Add "total_balls_held", m_balls_held
        End With
        DispatchPinEvent m_name & "_held_ball", kwargs

        'check if we are full now and post event if yes
        If IsFull() Then
            Set kwargs = GlfKwargs()
            With kwargs
                .Add "balls", m_balls_held
            End With
            DispatchPinEvent m_name & "_full", kwargs
        End If

        m_hold_queue.Add device, unclaimed_balls

        HoldBall = unclaimed_balls - balls_to_hold
    End Function

    Public Function ReleaseAll()
        'Release all balls in hold.
        ReleaseAll = ReleaseBalls(m_balls_held)
    End Function

    Public Function ReleaseBalls(balls_to_release)
        'Release all balls and return the actual amount of balls released.
        '
        'Args:
        '----
        '    balls_to_release: number of ball to release from hold
        
        If Ubound(m_hold_queue.Keys()) = -1 Then
            ReleaseBalls = 0
            Exit Function
        End If

        Dim remaining_balls_to_release : remaining_balls_to_release = balls_to_release

        m_base_device.Log "Releasing up to " & balls_to_release & " balls from hold"
        Dim balls_released : balls_released = 0
        Do While Ubound(m_hold_queue.Keys()) > -1
            Dim keys : keys = m_hold_queue.Keys()
            Dim device, balls_held
            device = keys(0)
            balls_held = m_hold_queue(device)
            m_hold_queue.Remove device

            Dim deviceControl : Set deviceControl = glf_ball_devices(device)
            
            Dim balls : balls = balls_held
            Dim balls_in_device : balls_in_device = deviceControl.Balls
            If balls > balls_in_device Then
                balls = balls_in_device
            End If

            If balls > remaining_balls_to_release Then
                m_hold_queue.Add device, balls_held - remaining_balls_to_release
                balls = remaining_balls_to_release
            End If

            deviceControl.EjectBalls balls
            balls_released = balls_released + balls
            remaining_balls_to_release = remaining_balls_to_release - balls
            If remaining_balls_to_release <= 0 Then
               Exit Do
            End If
        Loop

        If balls_released > 0 Then
            Dim kwargs : Set kwargs = GlfKwargs()
            With kwargs
                .Add "balls_released", balls_released
            End With
            DispatchPinEvent m_name & "_balls_released", kwargs
        End If

        m_balls_held = m_balls_held - balls_released
        ReleaseBalls = balls_released
    End Function

    Public Function ToYaml()
        Dim yaml
        Dim evt, x
        yaml = "  " & Replace(m_name, "ball_hold_", "") & ":" & vbCrLf
        If UBound(m_base_device.EnableEvents().Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents().keys
                yaml = yaml & m_base_device.EnableEvents()(key).Raw
                If x <> UBound(m_base_device.EnableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_base_device.DisableEvents().Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents().keys
                yaml = yaml & m_base_device.DisableEvents()(key).Raw
                If x <> UBound(m_base_device.DisableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        yaml = yaml & "    hold_devices: " & Join(m_hold_devices, ",") & vbCrLf
        If m_balls_to_hold > 0 Then
            yaml = yaml & "    balls_to_hold: " & m_balls_to_hold & vbCrLf
        End If
        If UBound(m_release_all_events.Keys) > -1 Then
            yaml = yaml & "    release_all_events: "
            x=0
            For Each key in m_release_all_events.keys
                yaml = yaml & m_release_all_events(key).Raw
                If x <> UBound(m_release_all_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_release_one_events.Keys) > -1 Then
            yaml = yaml & "    release_one_events: "
            x=0
            For Each key in m_release_one_events.keys
                yaml = yaml & m_release_one_events(key).Raw
                If x <> UBound(m_release_one_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_release_one_if_full_events.Keys) > -1 Then
            yaml = yaml & "    release_one_if_full_events: "
            x=0
            For Each key in m_release_one_if_full_events.keys
                yaml = yaml & m_release_one_if_full_events(key).Raw
                If x <> UBound(m_release_one_if_full_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function BallHoldsEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ball_hold : Set ball_hold = ownProps(1)
    Dim glfEvent
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If

    Select Case evt
        Case "hold"
            kwargs = ball_hold.HoldBall(ownProps(2), kwargs)
        Case "release_all"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            ball_hold.ReleaseAll
        Case "release_one"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            ball_hold.ReleaseBalls 1
        Case "release_one_if_full"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            If ball_hold.IsFull Then
                ball_hold.ReleaseBalls 1
            End If
    End Select

    If IsObject(args(1)) Then
        Set BallHoldsEventHandler = kwargs
    Else
        BallHoldsEventHandler = kwargs
    End If
End Function
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
    Public Property Let ActiveTime(value) : m_active_time = Glf_ParseInput(value) : End Property
    Public Property Let GracePeriod(value) : m_grace_period = Glf_ParseInput(value) : End Property
    Public Property Let HurryUpTime(value) : m_hurry_up_time = Glf_ParseInput(value) : End Property
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
        m_name = "ball_save_" & name
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
        RemoveDelay "_ball_save_"&m_name&"_disable"
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
        DispatchPinEvent m_name & "_hurry_up", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & Replace(m_name, "ball_save_", "") & ":" & vbCrLf
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
        yaml = yaml & "    auto_launch: " & LCase(m_auto_launch) & vbCrLf
        yaml = yaml & "    balls_to_save: " & m_balls_to_save & vbCrLf
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

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & m_events(key).Raw & ": " & vbCrLf
                For Each evt in m_eventValues(key)
                    yaml = yaml & "    - " & evt & vbCrLf
                Next
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function EventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            eventPlayer.FireEvent ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set EventPlayerEventHandler = kwargs
    Else
        EventPlayerEventHandler = kwargs
    End If
End Function

Class GlfLightPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = m_name : End Property
    
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get Events(name)
        If m_events.Exists(name) Then
            Set Events = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfLightPlayerEventItem)()
            m_events.Add name, new_event
            Set Events = new_event
        End If
    End Property

	Public default Function init(mode)
        m_name = "light_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_events = CreateObject("Scripting.Dictionary")
        
        Set m_base_device = (new GlfBaseModeDevice)(mode, "light_player", Me)
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

    Public Sub ReloadLights()
        m_base_device.Log "Reloading Lights"
        Dim evt
        For Each evt in m_events.Keys()
            Dim lightName, light
            'First get light counts
            For Each lightName in m_events(evt).LightNames
                Set light = m_events(evt).Lights(lightName)
                Dim lightsCount, x,tagLight, tagLights
                lightsCount = 0
                If Not glf_lightNames.Exists(lightName) Then
                    tagLights = glf_lightTags(lightName).Keys()
                    m_base_device.Log "Tag Lights: " & Join(tagLights)
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            Next
            m_base_device.Log "Adding " & lightsCount & " lights for event: " & evt 
            Dim seqArray
            ReDim seqArray(lightsCount-1)
            x=0
            'Build Seq
            For Each lightName in m_events(evt).LightNames
                Set light = m_events(evt).Lights(lightName)

                If Not glf_lightNames.Exists(lightName) Then
                    tagLights = glf_lightTags(lightName).Keys()
                    For Each tagLight in tagLights
                        seqArray(x) = tagLight & "|100|" & light.Color
                        x=x+1
                    Next
                Else
                    seqArray(x) = lightName & "|100|" & light.Color
                    x=x+1
                End If
            Next
            m_base_device.Log "Light List: " & Join(seqArray)
            m_events(evt).LightSeq = seqArray
        Next   
    End Sub

    Public Sub Play(evt, lights)
        LightPlayerCallbackHandler evt, Array(lights.LightSeq), m_name, m_priority
    End Sub

    Public Sub PlayOff(evt, lights)
        LightPlayerCallbackHandler evt, Null, m_name, m_priority
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function
End Class

Class GlfLightPlayerEventItem
	Private m_lights, m_lightSeq
  
    Public Property Get LightNames() : LightNames = m_lights.Keys() : End Property
    Public Property Get Lights(name)
        If m_lights.Exists(name) Then
            Set Lights = m_lights(name)
        Else
            Dim new_event : Set new_event = (new GlfLightPlayerItem)()
            m_lights.Add name, new_event
            Set Lights = new_event
        End If    
    End Property

    Public Property Get LightSeq() : LightSeq = m_lightSeq : End Property
    Public Property Let LightSeq(input) : m_lightSeq = input : End Property

    Public Function ToYaml()
        Dim yaml
        If UBound(m_lights.Keys) > -1 Then
            For Each key in m_lights.keys
                yaml = yaml & "    " & key & ": " & vbCrLf
                yaml = yaml & m_lights(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        Set m_lights = CreateObject("Scripting.Dictionary")
        m_lightSeq = Array()
        Set Init = Me
	End Function

End Class

Class GlfLightPlayerItem
	Private m_light, m_color, m_fade, m_priority
  
    Public Property Get Light(): Light = m_light: End Property
    Public Property Let Light(input): m_light = input: End Property

    Public Property Get Color(): Color = m_color: End Property
    Public Property Let Color(input): m_color = input: End Property

    Public Property Get Fade(): Fade = m_fade: End Property
    Public Property Let Fade(input): m_fade = input: End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input): m_priority = input: End Property      
  
    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "      color: " & m_color & vbCrLf
        If Not IsEmpty(m_fade) Then
            yaml = yaml & "      fade: " & m_fade & vbCrLf
        End If
        If m_priority <> 0 Then
            yaml = yaml & "      priority: " & m_priority & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        m_color = "ffffff"
        m_fade = Empty
        m_priority = 0
        Set Init = Me
	End Function

End Class

Function LightPlayerCallbackHandler(key, lights, mode, priority)
        
    Dim lightParts, light
    If IsNull(lights) Then
        
        glf_debugLog.WriteToLog "LightPlayer", "Removing Light Seq" & mode & "_" & key
    Else
        If UBound(lights) = -1 Then
            Exit Function
        End If
        If IsArray(lights) Then
            'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key    
        Else
            glf_debugLog.WriteToLog "LightPlayer", "Lights not an array!?"
        End If
        'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key
        Dim lightStack
        For Each light in lights(0)
            lightParts = Split(light,"|")
            
            Set lightStack = glf_lightStacks(lightParts(0))
            
            If lightStack.IsEmpty() Then
                ' If stack is empty, push the color onto the stack and set the light color
                lightStack.Push mode & "_" & key, lightParts(2), priority
                Glf_SetLight lightParts(0), lightParts(2)
            Else
                Dim current
                Set current = lightStack.Peek()
                If priority >= current("Priority") Then
                    ' If the new priority is higher, push it onto the stack and change the light color
                    lightStack.Push mode & "_" & key, lightParts(2), priority
                    Glf_SetLight lightParts(0), lightParts(2)
                Else
                    ' Otherwise, just push it onto the stack without changing the light color
                    lightStack.Push mode & "_" & key, lightParts(2), priority
                End If
            End If
        Next
        
        'TODO - Refactor this, this is the light fading. need to handle this differently
        'For Each light in lights(0)
        '    lightParts = Split(light, "|")
        '    If IsArray(lightParts) Then
        '        If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
        '           Dim tagLight, tagLights
        '            tagLights = lightCtrl.GetLightsForTag(lightParts(0))
        '            For Each tagLight in tagLights
        '                ProcessLight tagLight, lightParts, key
        '                If UBound(lightParts) = 2 Then
        '                    lightCtrl.FadeLightToColor tagLight, lightParts(1), lightParts(2), mode & "_" & key & "_" & tagLight, priority
        '                End If
        '            Next
        '        Else
        '            If UBound(lightParts) = 2 Then
        '                lightCtrl.FadeLightToColor lightParts(0), lightParts(1), lightParts(2), mode & "_" & key & "_" & lightParts(0), priority
        '            End If
        '        End If
        '    End If
        'Next
    End If
End Function

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

Class GlfLightStack
    Private stack

    Public default Function Init()
        ReDim stack(-1)  ' Initialize an empty array
        Set Init = Me
    End Function

    Public Sub Push(key, color, priority)
        Dim found : found = False
        Dim i

        ' Check if the key already exists in the stack and update it
        For i = LBound(stack) To UBound(stack)
            If stack(i)("Key") = key Then
                ' Replace the existing item if the key matches
                Set stack(i) = CreateColorPriorityObject(key, color, priority)
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            ' Insert the new item into the array maintaining priority order
            ReDim Preserve stack(UBound(stack) + 1)
            Set stack(UBound(stack)) = CreateColorPriorityObject(key, color, priority)
            SortStackByPriority
        End If
    End Sub

    ' Pop the top color from the stack
    Public Function Pop()
        If UBound(stack) >= 0 Then
            Set Pop = stack(UBound(stack))
            ReDim Preserve stack(UBound(stack) - 1)
        Else
            Set Pop = Nothing
        End If
    End Function

    ' Get the current top color without popping it
    Public Function Peek()
        If UBound(stack) >= 0 Then
            Set Peek = stack(UBound(stack))
        Else
            Set Peek = Nothing
        End If
    End Function

    ' Check if the stack is empty
    Public Function IsEmpty()
        IsEmpty = (UBound(stack) < 0)
    End Function

    ' Create a color-priority object
    Private Function CreateColorPriorityObject(key, color, priority)
        Dim colorPriorityObject
        Set colorPriorityObject = CreateObject("Scripting.Dictionary")
        colorPriorityObject.Add "Key", key
        colorPriorityObject.Add "Color", color
        colorPriorityObject.Add "Priority", priority
        Set CreateColorPriorityObject = colorPriorityObject
    End Function

    ' Sort the stack by priority (descending)
    Private Sub SortStackByPriority()
        Dim i, j
        Dim temp
        For i = LBound(stack) To UBound(stack) - 1
            For j = i + 1 To UBound(stack)
                If stack(i)("Priority") < stack(j)("Priority") Then
                    ' Swap the elements
                    Set temp = stack(i)
                    Set stack(i) = stack(j)
                    Set stack(j) = temp
                End If
            Next
        Next
    End Sub
End Class

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
            AddPinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable", "BaseModeDeviceEventHandler", m_priority, Array("enable", m_parent, m_enable_events(evt))
        Next
        For Each evt In m_disable_events.Keys()
            AddPinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable", "BaseModeDeviceEventHandler", m_priority, Array("disable", m_parent, m_disable_events(evt))
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
            glf_debugLog.WriteToLog m_mode.Name & m_device & "_" & m_parent.Name, message
        End If
    End Sub
End Class


Function BaseModeDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Dim glfEvent
    Select Case evt
        Case "activate"
            device.Activate
        Case "deactivate"
            device.Deactivate
        Case "enable"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            device.Enable
        Case "disable"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
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
    Private m_shot_profiles

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Get Debug(): Debug = m_debug: End Property
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
    End Property
    Public Property Let Debug(value)
        m_debug = value
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
        m_lightplayer = Null
        m_showplayer = Null
        m_segment_display_player = Null
        Set m_eventplayer = (new GlfEventPlayer)(Me)
        Set m_variableplayer = (new GlfVariablePlayer)(Me)
        Dim newEvent : Set newEvent = (new GlfEvent)("ball_ended")
        AddPinEventListener newEvent.EventName, m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me, newEvent)
        Set Init = Me
	End Function

    Public Sub StartMode()
        Log "Starting"
        m_started=True
        DispatchPinEvent m_name & "_starting", Null
        DispatchPinEvent m_name & "_started", Null
        Log "Started"
    End Sub

    Public Sub StopMode()
        If m_started = True Then
            m_started = False
            Log "Stopping"
            DispatchPinEvent m_name & "_stopping", Null
            DispatchPinEvent m_name & "_stopped", Null
            Log "Stopped"
        End If
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

Class GlfSegmentDisplayPlayer

    Private m_priority
    Private m_mode
    Private m_name
    Private m_events
    private m_base_device

    Public Property Get Name() : Name = "segment_player" : End Property

    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get Events(name)
        If m_events.Exists(name) Then
            Set Events = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfSegmentPlayerEventItem)()
            m_events.Add name, new_event
            Set Events = new_event
        End If
    End Property

	Public default Function init(mode)
        m_name = "segment_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "segment_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_segment_player_play", "SegmentPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_segment_player_play"
            PlayOff evt, m_events(evt)
        Next
    End Sub

    Public Sub Play(evt, segment_item)
        SegmentPlayerCallbackHandler evt, segment_item, m_mode, m_priority
    End Sub

    Public Sub PlayOff(evt, segment_item)
        Dim key
        key = m_mode & "." & "segment_player_player." & segment_item.Display
        If Not IsEmpty(segment_item.Key) Then
            key = key & segment_item.Key
        End If
        Dim display : Set display = glf_segment_displays(segment_item.Display)
        RemoveDelay key
        display.RemoveTextByKey key    
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfSegmentPlayerEventItem
	
    private m_display
    private m_text
    private m_priority
    private m_action
    private m_expire
    private m_flash_mask
    private m_flashing
    private m_key
    private m_transition
    private m_transition_out
    private m_color

    Public Property Get Display() : Display = m_display : End Property
    Public Property Let Display(input) : m_display = input : End Property
    
    Public Property Get Text()
        If Not IsNull(m_text) Then
            Text = m_text.Value()
        Else
            Text = Empty
        End If
    End Property
    Public Property Let Text(input) 
        Set m_text = (new GlfInput)(input)
    End Property

    Public Property Get Priority() : Priority = m_priority : End Property
    Public Property Let Priority(input) : m_priority = input : End Property
        
    Public Property Get Action() : Action = m_action : End Property
    Public Property Let Action(input) : m_action = input : End Property
                
    Public Property Get Expire() : Expire = m_expire : End Property
    Public Property Let Expire(input) : m_events = input : End Property

    Public Property Get FlashMask() : FlashMask = m_flash_mask : End Property
    Public Property Let FlashMask(input) : m_flash_mask = input : End Property
                        
    Public Property Get Flashing() : Flashing = m_flashing : End Property
    Public Property Let Flashing(input) : m_flashing = input : End Property
                            
    Public Property Get Key() : Key = m_key : End Property
    Public Property Let Key(input) : m_key = input : End Property

    Public Property Get Color() : Color = m_color : End Property
    Public Property Let Color(input) : m_color = input : End Property

    Public Property Get Transition()
        If IsNull(m_transition) Then
            Set m_transition = (new GlfSegmentPlayerTransition)()
            Set Transition = m_transition   
        Else
            Set Transition = m_transition
        End If
    End Property

    Public Property Get TransitionOut()
        If IsNull(m_transition_out) Then
            Set m_transition_out = (new GlfSegmentPlayerTransition)()
            Set TransitionOut = m_transition_out   
        Else
            Set TransitionOut = m_transition_out
        End If
    End Property
                                
	Public default Function init()
        m_display = Empty
        m_text = Null
        m_priority = 0
        m_action = "add"
        m_expire = 0
        m_flash_mask = Empty
        m_flashing = "not_set"
        m_key = Empty
        m_transition = Null
        m_transition_out = Null
        m_color = Rgb(255,255,255)
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        If Not IsEmpty(m_display) Then
            yaml = yaml & "    " & m_display & ": " & vbCrLf
        End If
        If Not IsNull(m_text) Then
            yaml = yaml & "    " & m_text.Raw() & ": " & vbCrLf
        End If
        If m_priority > 0 Then
            yaml = yaml & "    " & m_priority & ": " & vbCrLf
        End If
        If m_action <> "add" Then
            yaml = yaml & "    " & m_action & ": " & vbCrLf
        End If
        If m_expire > 0 Then
            yaml = yaml & "    " & m_expire & ": " & vbCrLf
        End If
        If Not IsEmpty(m_flash_mask) Then
            yaml = yaml & "    " & m_flash_mask & ": " & vbCrLf
        End If
        If m_flashing <> "not_set" Then
            yaml = yaml & "    " & m_flashing & ": " & vbCrLf
        End If
        If Not IsEmpty(m_key) Then
            yaml = yaml & "    " & m_key & ": " & vbCrLf
        End If
        If Not IsEmpty(m_color) Then
            yaml = yaml & "    " & m_color & ": " & vbCrLf
        End If
        If Not IsNull(m_transition) Then
            yaml = yaml & m_transition.ToYaml()
        End If
        If Not IsNull(m_transition_out) Then
            yaml = yaml & m_transition_out.ToYaml()
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfSegmentPlayerTransition
	
    private m_type
    private m_text
    private m_direction

    Public Property Get TransitionType() : TransitionType = m_type : End Property
    Public Property Let TransitionType(input) : m_type = input : End Property
    
    Public Property Get Text() : Text = m_text : End Property
    Public Property Let Text(input) : m_text = input : End Property

    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let Direction(input) : m_direction = input : End Property                          

	Public default Function init()
        m_type = "push"
        m_text = Empty
        m_direction = "push"
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    transition:" & vbCrLf
        yaml = yaml & "      " & m_type & ": " & vbCrLf
        yaml = yaml & "      " & m_direction & ": " & vbCrLf
        yaml = yaml & "      " & m_text & ": " & vbCrLf
        ToYaml = yaml
    End Function

End Class

Function SegmentPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim SegmentPlayer : Set SegmentPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            SegmentPlayer.Activate
        Case "deactivate"
            SegmentPlayer.Deactivate
        Case "play"
            SegmentPlayer.Play ownProps(3), ownProps(2)
    End Select
    SegmentPlayerEventHandler = Null
End Function


Function SegmentPlayerCallbackHandler(evt, segment_item, mode, priority)

    If IsObject(segment_item) Then
        'Shot Text on a display
        Dim key
        key = mode & "." & "segment_player_player." & segment_item.Display
        
        If Not IsEmpty(segment_item.Key) Then
            key = key & segment_item.Key
        End If

        Dim display : Set display = glf_segment_displays(segment_item.Display)
        
        If segment_item.Action = "add" Then
            RemoveDelay key
            display.AddTextEntry segment_item.Text, segment_item.Color, segment_item.Flashing, segment_item.FlashMask, segment_item.Transition, segment_item.TransitionOut, segment_item.Priority, segment_item.Key
                                
            If segment_item.Expire > 0 Then
                'TODO Add delay for remove
                'SetDelay
            End If

        ElseIf segment_item.Action = "remove" Then
            RemoveDelay key
            display.RemoveTextByKey key        
        ElseIf segment_item.Action = "flash" Then
            display.SetFlashing "all"
        ElseIf segment_item.Action = "flash_match" Then
            display.SetFlashing "match"
        ElseIf segment_item.Action = "flash_mask" Then
            display.SetFlashingMask segment_item.FlashMask
        ElseIf segment_item.Action = "no_flash" Then
            display.SetFlashing "no_flash"
        ElseIf segment_item.Action = "set_color" Then
            If Not IsNull(segment_item.Color) Then
                display.SetColor segment_item.Color
            End If
        End If
    End If

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
    Private m_isRotating
 
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
        m_isRotating = False
 
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
        Dim evt
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
        Dim shot_name
        For Each shot_name in m_shots
            AddPinEventListener shot_name & "_hit", m_name & "_" & m_mode & "_hit", "ShotGroupEventHandler", m_priority, Array("hit", Me)
            AddPlayerStateEventListener "player_shot_" & shot_name, m_name & "_" & m_mode & "_complete", "ShotGroupEventHandler", m_priority, Array("complete", Me)
        Next
    End Sub
 
    Public Sub Deactivate
        Dim evt
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
        Dim shot_name
        For Each shot_name in m_shots
            RemovePinEventListener shot_name & "_hit", m_name & "_" & m_mode & "_hit"
            RemovePlayerStateEventListener "player_shot_" & shot_name, m_name & "_" & m_mode & "_complete"
        Next
    End Sub
 
    Public Function CheckForComplete()
        If m_isRotating Then
            Exit Function
        End If
        Dim state : state = CommonState()
        If state = m_common_state Then
            Exit Function
        End If
 
        m_common_state = state
 
        If state = Empty Then
            Exit Function
        End If
 
        Dim state_name : state_name = Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).StateName(m_common_state)
 
        m_base_device.Log "Shot group is complete with state: " & state_name
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
        m_isRotating = True
        For Each shot in m_temp_shots.Keys
            m_base_device.Log "Rotating Shot:" & shot
            m_temp_shots(shot).Jump shot_states(x), True, False
            x=x+1
        Next 
        m_isRotating = False
        CheckForComplete()
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
        Case "complete"
            device.CheckForComplete
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

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "  " & Replace(m_name, "shotprofile_", "") & ":" & vbCrLf
        yaml = yaml & "    states: " & vbCrLf
        Dim token,evt,state,x : x = 0
        For Each evt in m_states.Keys
            Set state = StateForIndex(x)
            yaml = yaml & "     - name: " & StateName(x) & vbCrLf
            yaml = yaml & "       show: " & state.Show.Name & vbCrLf
            yaml = yaml & "       loops: " & m_states(evt).Loops & vbCrLf
            yaml = yaml & "       speed: " & m_states(evt).Speed & vbCrLf
            yaml = yaml & "       sync_ms: " & m_states(evt).SyncMs & vbCrLf

            If Ubound(state.Tokens().Keys)>-1 Then
                yaml = yaml & "       show_tokens: " & vbCrLf
                For Each token in state.Tokens().Keys()
                    yaml = yaml & "         " & token & ": " & state.Tokens()(token) & vbCrLf
                Next
            End If

            'yaml = yaml & "     block: " & m_block & vbCrLf
            'yaml = yaml & "     advance_on_hit: " & m_advance_on_hit & vbCrLf
            'yaml = yaml & "     loop: " & m_loop & vbCrLf
            'yaml = yaml & "     rotation_pattern: " & m_rotation_pattern & vbCrLf
            'yaml = yaml & "     state_names_to_not_rotate: " & m_states_not_to_rotate & vbCrLf
            'yaml = yaml & "     state_names_to_rotate: " & m_states_to_rotate & vbCrLf
            x = x +1
        Next
        ToYaml = yaml
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
    Private m_player_var_name
    Private m_persist
    Private m_internal_cache_id

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
    
    Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property

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
    Public Property Let Persist(value) : m_persist = value : End Property
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
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_internal_cache_id = -1
        m_enabled = False
        m_persist = True
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot", Me)

        m_profile = "default"
        m_player_var_name = "player_shot_" & m_name
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
        If IsNull(GetPlayerState(m_player_var_name)) Then
            m_state = 0
            If m_persist Then
                SetPlayerState m_player_var_name, 0
            End If
        Else
            m_state = GetPlayerState(m_player_var_name)
        End If
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
        For Each evt in m_hit_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        Dim x
        For x=0 to Glf_ShotProfiles(m_profile).StatesCount()
            StopShowForState(x)
        Next
    End Sub

    Private Sub StopShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        m_base_device.Log "Removing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        If glf_running_shows.Exists(m_mode & "_" & CStr(state) & "_" & m_name & "_" & profileState.Key) Then 
            glf_running_shows(m_mode & "_" & CStr(state) & "_" & m_name & "_" & profileState.Key).StopRunningShow()
        End If
    End Sub

    Private Sub PlayShowForState(state)
        If m_enabled = False Then
            Exit Sub
        End If
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        m_base_device.Log "Playing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        If IsObject(profileState) Then
            If Not IsNull(profileState.Show) Then
                Dim new_running_show
                Set new_running_show = (new GlfRunningShow)(m_mode & "_" & CStr(m_state) & "_" & m_name & "_" & profileState.Key, profileState.Key, profileState, m_priority + profileState.Priority, m_tokens, m_internal_cache_id)
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
                If m_persist Then
                    SetPlayerState m_player_var_name, 0
                End If
                PlayShowForState(m_state)
            Else
                Advance = False
                Exit Function
            End If
        Else
            StopShowForState(m_state)
            m_state = m_state + 1
            If m_persist Then
                SetPlayerState m_player_var_name, m_state
            End If
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
        If m_persist Then
            SetPlayerState m_player_var_name, m_state
        End If
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
        ElseIf UBound(m_switches) > 0 Then
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

        If UBound(m_base_device.EnableEvents().Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents().keys
                yaml = yaml & m_base_device.EnableEvents()(key).Raw
                If x <> UBound(m_base_device.EnableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_base_device.DisableEvents().Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents().keys
                yaml = yaml & m_base_device.DisableEvents()(key).Raw
                If x <> UBound(m_base_device.DisableEvents().Keys) Then
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
                yaml = yaml & "        state: " & m_control_events(key).State & vbCrLf
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
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "show_player" : End Property
    Public Property Get EventShows() : EventShows = m_events.Items() : End Property
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
        Set m_base_device = (new GlfBaseModeDevice)(mode, "show_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            m_base_device.Log "Adding EVENT: " & Replace(evt, "_" & m_events(evt).Key, "")
            AddPinEventListener evt , m_mode & "_" & m_events(evt).Key & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_" & m_events(evt).Key & "_show_player_play"
            PlayOff m_events(evt).Key
        Next
    End Sub

    Public Sub Play(evt, show)
        If show.Action = "stop" Then
           PlayOff show.Key
        Else
            Dim new_running_show
            Set new_running_show = (new GlfRunningShow)(m_name & "_" & show.Key, show.Key, show, m_priority, Null, Null)
        End If
    End Sub

    Public Sub PlayOff(key)
        If glf_running_shows.Exists(m_name & "_" & key) Then 
            glf_running_shows(m_name & "_" & key).StopRunningShow()
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

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
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms, m_duration, m_priority, m_internal_cache_id
  
	Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property

    Public Property Get Show()
        If IsNull(m_show) Then
            Show = Null
        Else
            Set Show = m_show
        End If
    End Property
	Public Property Let Show(input): Set m_show = input: End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property        

    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property        
  
    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    " & m_show.Name &": " & vbCrLf
        If m_action <> "play" Then
            yaml = yaml & "      action: " & m_action & vbCrLf
        End If
        If m_key <> "" Then
            yaml = yaml & "      key: " & m_key & vbCrLf
        End If
        If m_priority <> 0 Then
            yaml = yaml & "      priority: " & m_priority & vbCrLf
        End If
        If m_loops > -1 Then
            yaml = yaml & "      loops: " & m_loops & vbCrLf
        End If
        If m_speed <> 1 Then
            yaml = yaml & "      speed: " & m_speed & vbCrLf
        End If
        If UBound(m_tokens.Keys) > -1 Then
            yaml = yaml & "      show_tokens: " & vbCrLf
            Dim key
            For Each key in m_tokens.Keys
                yaml = yaml & "        " & key & ": " & m_tokens(key) & vbCrLf
            Next
        End If
        If m_syncms > 0 Then
            yaml = yaml & "      sync_ms: " & m_syncms & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        m_action = "play"
        m_key = ""
        m_priority = 0
        m_loops = -1
        m_internal_cache_id = -1
        m_speed = 1
        m_syncms = 0
        m_show = Null
        Set m_tokens = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class


Class GlfShow

    Private m_name
    Private m_steps
    Private m_total_step_time

    Public Property Get Name(): Name = m_name: End Property
    
    Public Property Get Steps() : Set Steps = m_steps : End Property

    Public Function StepAtIndex(index) : Set StepAtIndex = m_steps.Items()(index) : End Function
    
    Public default Function init(name)
        m_name = name
        m_total_step_time = 0
        Set m_steps = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

    Public Function AddStep(absolute_time, relative_time, duration)
        Dim new_step : Set new_step = (new GlfShowStep)()
        new_step.Duration = duration
        new_step.RelativeTime = relative_time
        new_step.AbsoluteTime = absolute_time
        new_step.IsLastStep = True
        
        'Add a empty first step if if show does not start right away
        If UBound(m_steps.Keys) = -1 Then
            If Not IsNull(new_step.Time) And new_step.Time <> 0 Then
                Dim empty_step : Set empty_step = (new GlfShowStep)()
                empty_step.Duration = new_step.Time
                m_total_step_time = new_step.Time
                m_steps.Add CStr(UBound(m_steps.Keys())+1), empty_step        
            End If
        End If
        

        

        If UBound(m_steps.Keys()) > -1 Then
            Dim prevStep : Set prevStep = m_steps.Items()(UBound(m_steps.Keys()))
            prevStep.IsLastStep = False
            'need to work out previous steps duration.
            If IsNull(prevStep.Duration) Then
                'The previous steps duration needs calculating.
                'If this step has a relative time then the last steps duration is that time.
                If Not IsNull(new_step.Time) Then
                    If new_step.IsRelativeTime Then
                        prevStep.Duration = new_step.Time
                    Else
                        prevStep.Duration = new_step.Time - m_total_step_time
                    End If
                Else
                    prevStep.Duration = 1
                End If
            End If
            m_total_step_time = m_total_step_time + prevStep.Duration
        Else
            If IsNull(new_step.Duration) Then
                m_total_step_time = m_total_step_time + 1
            Else
                m_total_step_time = m_total_step_time + new_step.Time
            End If
        End If

        m_steps.Add CStr(UBound(m_steps.Keys())+1), new_step
        Set AddStep = new_step
    End Function

    Public Function ToYaml()
        'Dim yaml
        'yaml = yaml & "  " & Replace(m_name, "shotprofile_", "") & ":" & vbCrLf
        'yaml = yaml & "    states: " & vbCrLf
        'Dim token,evt,x : x = 0
        'For Each evt in m_states.Keys
        '    yaml = yaml & "     - name: " & StateName(x) & vbCrLf
            'yaml = yaml & "       show: " & m_states(evt).Show & vbCrLf
            'yaml = yaml & "       loops: " & m_states(evt).Loops & vbCrLf
            'yaml = yaml & "       speed: " & m_states(evt).Speed & vbCrLf
            'yaml = yaml & "       sync_ms: " & m_states(evt).SyncMs & vbCrLf

            'If Ubound(m_states(evt).Tokens().Keys)>-1 Then
            '    yaml = yaml & "       show_tokens: " & vbCrLf
            '    For Each token in m_states(evt).Tokens().Keys()
            '        yaml = yaml & "         " & token & ": " & m_states(evt).Tokens(token) & vbCrLf
            '    Next
            'End If

            'yaml = yaml & "     block: " & m_block & vbCrLf
            'yaml = yaml & "     advance_on_hit: " & m_advance_on_hit & vbCrLf
            'yaml = yaml & "     loop: " & m_loop & vbCrLf
            'yaml = yaml & "     rotation_pattern: " & m_rotation_pattern & vbCrLf
            'yaml = yaml & "     state_names_to_not_rotate: " & m_states_not_to_rotate & vbCrLf
            'yaml = yaml & "     state_names_to_rotate: " & m_states_to_rotate & vbCrLf
         '   x = x +1
        'Next
        'ToYaml = yaml
    End Function

End Class

Class GlfRunningShow

    Private m_key
    Private m_show_name
    Private m_show_settings
    Private m_current_step
    Private m_priority
    Private m_total_steps
    Private m_tokens
    Private m_internal_cache_id

    Public Property Get CacheName(): CacheName = m_show_name & "_" & m_internal_cache_id & "_" & ShowSettings.InternalCacheId: End Property
    Public Property Get Tokens(): Set Tokens = m_tokens : End Property

    Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property        

    Public Property Get CurrentStep(): CurrentStep = m_current_step End Property
    Public Property Let CurrentStep(input): m_current_step = input End Property        

    Public Property Get TotalSteps(): TotalSteps = m_total_steps End Property
    Public Property Let TotalSteps(input): m_total_steps = input End Property        

    Public Property Get ShowName(): ShowName = m_show_name: End Property
    Public Property Let ShowName(input): m_show_name = input: End Property
        
    Public Property Get ShowSettings(): Set ShowSettings = m_show_settings: End Property
    Public Property Let ShowSettings(input): Set m_show_settings = input: End Property
    
    Public default Function init(rname, rkey, show_settings, priority, tokens, cache_id)
        m_show_name = rname
        m_key = rkey
        m_current_step = 0
        m_priority = priority
        m_internal_cache_id = cache_id
        Set m_show_settings = show_settings

        Dim key
        Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
        If Not IsNull(m_show_settings.Tokens) Then
            For Each key In m_show_settings.Tokens.Keys()
                mergedTokens.Add key, m_show_settings.Tokens()(key)
            Next
        End If
        If Not IsNull(tokens) Then
            For Each key In tokens.Keys
                If mergedTokens.Exists(key) Then
                    mergedTokens(key) = tokens(key)
                Else
                    mergedTokens.Add key, tokens(key)
                End If
            Next
        End If
        Set m_tokens = mergedTokens
        
        m_total_steps = UBound(m_show_settings.Show.Steps.Keys())
        If glf_running_shows.Exists(m_show_name) Then
            glf_running_shows(m_show_name).StopRunningShow()
            glf_running_shows.Add m_show_name, Me
        Else
            glf_running_shows.Add m_show_name, Me
        End If 
        Play
        Set Init = Me
	End Function

    Public Sub Play()
        'Play the show.
        Log "Playing show: " & m_show_name & " With Key: " & m_key
        GlfShowStepHandler(Array(Me))
    End Sub

    Public Sub StopRunningShow()
        Log "Removing show: " & m_show_name & " With Key: " & m_key
        Dim cached_show,light, cached_show_lights
        If glf_cached_shows.Exists(CacheName) Then
            cached_show = glf_cached_shows(CacheName)
            Set cached_show_lights = cached_show(1)
        Else
            msgbox "show not cached! Problem with caching"
        End If
        Dim lightStack
        For Each light in cached_show_lights.Keys()

            Set lightStack = glf_lightStacks(light)
            
            If Not lightStack.IsEmpty() Then
                ' Pop the current top color
                lightStack.Pop()
            End If
            
            If Not lightStack.IsEmpty() Then
                ' Set the light to the next color on the stack
                Dim nextColor
                Set nextColor = lightStack.Peek()
                Glf_SetLight light, nextColor("Color")
            Else
                ' Turn off the light since there's nothing on the stack
                Glf_SetLight light, "000000"
            End If
        Next

        RemoveDelay Me.ShowName & "_" & Me.Key
        glf_running_shows.Remove m_show_name
    End Sub

    Public Sub Log(message)
        glf_debugLog.WriteToLog "Running Show", message
    End Sub
End Class

Function GlfShowStepHandler(args)
    Dim running_show : Set running_show = args(0)
    Dim nextStep : Set nextStep = running_show.ShowSettings.Show.StepAtIndex(running_show.CurrentStep)
    If UBound(nextStep.Lights) > -1 Then
        Dim cached_show, cached_show_seq
        If glf_cached_shows.Exists(running_show.CacheName) Then
            cached_show = glf_cached_shows(running_show.CacheName)
            cached_show_seq = cached_show(0)
        Else
            msgbox "show not cached! Problem with caching"
        End If
'        glf_debugLog.WriteToLog "Running Show", join(cached_show(running_show.CurrentStep))
        LightPlayerCallbackHandler running_show.Key, Array(cached_show_seq(running_show.CurrentStep)), running_show.ShowName, running_show.Priority + running_show.ShowSettings.Priority
    End If
    If nextStep.Duration = -1 Then
        'glf_debugLog.WriteToLog "Running Show", "HOLD"
        Exit Function
    End If
    running_show.CurrentStep = running_show.CurrentStep + 1
    If nextStep.IsLastStep = True Then
        'msgbox "last step"
        If IsNull(nextStep.Duration) Then
            'msgbox "5!"
            nextStep.Duration = 1
        End If
    End If
    If running_show.CurrentStep > running_show.TotalSteps Then
        'End of Show
        'glf_debugLog.WriteToLog "Running Show", "END OF SHOW"
        If running_show.ShowSettings.Loops = -1 Or running_show.ShowSettings.Loops > 1 Then
            If running_show.ShowSettings.Loops > 1 Then
                running_show.ShowSettings.Loops = running_show.ShowSettings.Loops - 1
            End If
            running_show.CurrentStep = 0
            SetDelay running_show.ShowName & "_" & running_show.Key, "GlfShowStepHandler", Array(running_show), (nextStep.Duration / running_show.ShowSettings.Speed) * 1000
        Else
'            glf_debugLog.WriteToLog "Running Show", "STOPPING SHOW, NO Loops"
            running_show.StopRunningShow()
        End If
    Else
'        glf_debugLog.WriteToLog "Running Show", "Scheduling Next Step"
        SetDelay running_show.ShowName & "_" & running_show.Key, "GlfShowStepHandler", Array(running_show), (nextStep.Duration / running_show.ShowSettings.Speed) * 1000
    End If
End Function

Class GlfShowStep

    Private m_lights, m_time, m_duration, m_isLastStep, m_absTime, m_relTime

    Public Property Get Lights(): Lights = m_lights: End Property
    Public Property Let Lights(input) : m_lights = input: End Property

    Public Property Get Time()
        If IsNull(m_relTime) Then
            Time = m_absTime
        Else
            Time = m_relTime
        End If
    End Property

    Public Property Get IsRelativeTime()
        If Not IsNull(m_relTime) Then
            IsRelativeTime = True
        Else
            IsRelativeTime = False
        End If
    End Property

    Public Property Let RelativeTime(input) : m_relTime = input: End Property
    Public Property Let AbsoluteTime(input) : m_absTime = input: End Property

    Public Property Get Duration(): Duration = m_duration: End Property
    Public Property Let Duration(input) : m_duration = input: End Property

    Public Property Get IsLastStep(): IsLastStep = m_isLastStep: End Property
    Public Property Let IsLastStep(input) : m_isLastStep = input: End Property
        
    Public default Function init()
        m_lights = Array()
        m_duration = Null
        m_time = Null
        m_absTime = Null
        m_relTime = Null
        m_isLastStep = False
        Set Init = Me
	End Function

    Public Function ToYaml()
        'Dim yaml
        'yaml = yaml & "  " & Replace(m_name, "shotprofile_", "") & ":" & vbCrLf
        'yaml = yaml & "    states: " & vbCrLf
        'Dim token,evt,x : x = 0
        'For Each evt in m_states.Keys
        '    yaml = yaml & "     - name: " & StateName(x) & vbCrLf
            'yaml = yaml & "       show: " & m_states(evt).Show & vbCrLf
            'yaml = yaml & "       loops: " & m_states(evt).Loops & vbCrLf
            'yaml = yaml & "       speed: " & m_states(evt).Speed & vbCrLf
            'yaml = yaml & "       sync_ms: " & m_states(evt).SyncMs & vbCrLf

            'If Ubound(m_states(evt).Tokens().Keys)>-1 Then
            '    yaml = yaml & "       show_tokens: " & vbCrLf
            '    For Each token in m_states(evt).Tokens().Keys()
            '        yaml = yaml & "         " & token & ": " & m_states(evt).Tokens(token) & vbCrLf
            '    Next
            'End If

            'yaml = yaml & "     block: " & m_block & vbCrLf
            'yaml = yaml & "     advance_on_hit: " & m_advance_on_hit & vbCrLf
            'yaml = yaml & "     loop: " & m_loop & vbCrLf
            'yaml = yaml & "     rotation_pattern: " & m_rotation_pattern & vbCrLf
            'yaml = yaml & "     state_names_to_not_rotate: " & m_states_not_to_rotate & vbCrLf
            'yaml = yaml & "     state_names_to_rotate: " & m_states_to_rotate & vbCrLf
         '   x = x +1
        'Next
        'ToYaml = yaml
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
    Private m_base_device

    Private m_control_events
    Private m_running
    Private m_ticks
    Private m_ticks_remaining
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_tick_interval
    Private m_starting_tick_interval
    Private m_max_value
    Private restart_on_complete
    Private m_start_running

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get ControlEvents(name)
        If m_control_events.Exists(name) Then
            Set ControlEvents = m_control_events(name)
        Else
            Dim newEvent : Set newEvent = (new GlfTimerControlEvent)()
            m_control_events.Add name, newEvent
            Set ControlEvents = newEvent
        End If
    End Property
    Public Property Get StartValue() : StartValue = m_start_value : End Property
    Public Property Get EndValue() : EndValue = m_end_value : End Property
    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let StartValue(value) : m_start_value = value : End Property
    Public Property Let EndValue(value) : m_end_value = value : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let MaxValue(value) : m_max_value = value : End Property
    Public Property Let RestartOnComplete(value) : restart_on_complete = value : End Property
    Public Property Let StartRunning(value) : m_start_running = value : End Property
    Public Property Let TickInterval(value)
        m_tick_interval = value * 1000
        m_starting_tick_interval = value
    End Property

    Public Property Get GetValue(value)
        Select Case value
            Case "ticks_remaining"
              GetValue = m_ticks_remaining
        End Select
    End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_direction = "up"
        m_ticks = 0
        m_ticks_remaining = 0
        m_tick_interval = 1000
        m_starting_tick_interval = 1
        restart_on_complete = False
        start_running = False

        Set m_control_events = CreateObject("Scripting.Dictionary")
        m_running = False

        Set m_base_device = (new GlfBaseModeDevice)(mode, "timer", Me)

        glf_timers.Add name, Me

        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_control_events.Keys
            AddPinEventListener m_control_events(evt).EventName, m_name & "_action", "TimerEventHandler", m_priority, Array("action", Me, m_control_events(evt))
        Next
        m_ticks = m_start_value
        m_ticks_remaining = m_ticks
        If m_start_running Then
            StartTimer()
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_control_events.Keys
            RemovePinEventListener m_control_events(evt).EventName, m_name & "_action"
        Next
        RemoveDelay m_name & "_tick"
        m_running = False
    End Sub

    Public Sub Action(controlEvent)

        dim value : value = controlEvent.Value
        Select Case controlEvent.Action
            Case "add"
                Add GetRef(value(0))()
            Case "subtract"
                Subtract GetRef(value(0))()
            Case "jump"
                Jump GetRef(value(0))()
            Case "start"
                StartTimer()
            Case "stop"
                StopTimer()
            Case "reset"
                Reset()
            Case "restart"
                Restart()
            Case "pause"
                Pause GetRef(value(0))()
            Case "set_tick_interval"
                SetTickInterval GetRef(value(0))()
            Case "change_tick_interval"
                ChangeTickInterval GetRef(value(0))()
            Case "reset_tick_interval"
                SetTickInterval m_starting_tick_interval
        End Select

    End Sub

    Private Sub StartTimer()
        If m_running Then
            Exit Sub
        End If

        m_base_device.Log "Starting Timer"
        m_running = True
        RemoveDelay m_name & "_unpause"
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_started", kwargs
        PostTickEvents()
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval
    End Sub

    Private Sub StopTimer()
        m_base_device.Log "Stopping Timer"
        m_running = False
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_stopped", kwargs
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub Pause(pause_ms)
        m_base_device.Log "Pausing Timer for "&pause_ms&" ms"
        m_running = False
        RemoveDelay m_name & "_tick"
        
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_paused", kwargs

        If pause_ms > 0 Then
            Dim startControlEvent : Set startControlEvent = (new GlfTimerControlEvent)()
            startControlEvent.Action = "start"
            SetDelay m_name & "_unpause", "TimerEventHandler", Array(Array("action", Me, startControlEvent), Null), pause_ms
        End If
    End Sub 

    Public Sub Tick()
        m_base_device.Log "Timer Tick"
        If Not m_running Then
            m_base_device.Log "Timer is not running. Will remove."
            Exit Sub
        End If

        Dim newValue
        If m_direction = "down" Then
            newValue = m_ticks - 1
        Else
            newValue = m_ticks + 1
        End If
        
        m_base_device.Log "ticking: old value: "& m_ticks & ", new Value: " & newValue & ", target: "& m_end_value
        m_ticks = newValue
        If Not PostTickEvents() Then
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval    
        End If
    End Sub

    Private Function CheckForDone

        ' Checks to see if this timer is done. Automatically called anytime the
        ' timer's value changes.
        m_base_device.Log "Checking to see if timer is done. Ticks: "&m_ticks&", End Value: "&m_end_value&", Direction: "& m_direction

        if m_direction = "up" And Not IsEmpty(m_end_value) And m_ticks >= m_end_value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If m_direction = "down" And m_ticks <= m_end_value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If Not IsEmpty(m_end_value) Then 
            m_ticks_remaining = abs(m_end_value - m_ticks)
        End If
        m_base_device.Log "Timer is not done"

        CheckForDone = False

    End Function

    Private Sub TimerComplete

        m_base_device.Log "Timer Complete"

        StopTimer()
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        
        If m_restart_on_complete Then
            m_base_device.Log "Restart on complete: True"
            Restart()
        End If
    End Sub

    Private Sub Restart
        Reset()
        If Not m_running Then
            StartTimer()
        Else
            PostTickEvents()
        End If
    End Sub

    Private Sub Reset
        m_base_device.Log "Resetting timer. New value: "& m_start_value
        Jump m_start_value
    End Sub

    Private Sub Jump(timer_value)
        m_ticks = (timer_value/1000)

        If m_max_value and m_ticks > m_max_value Then
            m_ticks = m_max_value
        End If

        CheckForDone()
    End Sub

    Public Sub ChangeTickInterval(change)
        m_tick_interval = m_tick_interval * (change/1000)
    End Sub

    Public Sub SetTickInterval(timer_value)
        m_tick_interval = timer_value
    End Sub
        
    Private Function PostTickEvents()
        PostTickEvents = True
        If Not CheckForDone() Then
            PostTickEvents = False
            Dim kwargs : Set kwargs = GlfKwargs()
            With kwargs
                .Add "ticks", m_ticks
                .Add "ticks_remaining", m_ticks_remaining
            End With
            DispatchPinEvent m_name & "_tick", kwargs
            m_base_device.Log "Ticks: "&m_ticks&", Remaining: " & m_ticks_remaining
        End If
    End Function

    Public Sub Add(timer_value) 
        Dim new_value

        new_value = m_ticks + (timer_value/1000)

        If Not IsEmpty(m_max_value) And new_value > m_max_value Then
            new_value = m_max_value
        End If
        m_ticks = new_value
        timer_value = new_value - timer_value

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_added", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_added", kwargs
        CheckForDone()
    End Sub

    Public Sub Subtract(timer_value)
        m_ticks = m_ticks - (timer_value/1000)
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_subtracted", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_subtracted", kwargs
        
        CheckForDone()
    End Sub
End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    
    Select Case evt
        Case "action"
            Dim controlEvent : Set controlEvent = ownProps(2)
            timer.Action controlEvent
        Case "tick"
            timer.Tick 
    End Select
    TimerEventHandler = kwargs
End Function

Class GlfTimerControlEvent
	Private m_event, m_action, m_value
  
	Public Property Get EventName(): EventName = m_event: End Property
    Public Property Let EventName(input): m_event = input: End Property

    Public Property Get Action(): Action = m_action : End Property
    Public Property Let Action(input): m_action = input : End Property

    Public Property Get Value()
        Value = m_value
    End Property
    Public Property Let Value(input)
        m_value = Glf_ParseInput(input)
    End Property

	Public default Function init()
        m_event = Empty
        m_action = Empty
        m_value = Empty
	    Set Init = Me
	End Function

End Class
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

	Public Property Let Float(input): m_float = Glf_ParseInput(input): m_type = "float" : End Property
  
	Public Property Let Int(input): m_int = Glf_ParseInput(input): m_type = "int" : End Property
  
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
    executionTime = gametime + delayInMs
    'If gametime >= executionTime Then
    '    executionTime = executionTime + 100
    'End If
    
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

    'glf_debugLog.WriteToLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc & ", ExecutionTime: " & executionTime
    delayQueue(executionTime).Add name, (new DelayObject)(name, callbackFunc, executionTime, args)
    delayQueueMap.Add name, executionTime
    
End Sub

Function AlignToNearest10th(timeMs)
    AlignToNearest10th = Int(timeMs / 100) * 100
End Function

Function RemoveDelay(name)
    If delayQueueMap.Exists(name) Then
        If delayQueue.Exists(delayQueueMap(name)) Then
            If delayQueue(delayQueueMap(name)).Exists(name) Then
                'glf_debugLog.WriteToLog "Delay", "Removing delay for " & name & " and  Execution Time: " & delayQueueMap(name)
                delayQueue(delayQueueMap(name)).Remove name
            End If
            delayQueueMap.Remove name
            RemoveDelay = True
            'glf_debugLog.WriteToLog "Delay", "Removing delay for " & name
            Exit Function
        End If
    End If
    RemoveDelay = False
End Function

Sub DelayTick()
    Dim queueItem, key, delayObject
    For Each queueItem in delayQueue.Keys()
        If Int(queueItem) < gametime Then
            For Each key In delayQueue(queueItem).Keys()
                If IsObject(delayQueue(queueItem)(key)) Then
                    Set delayObject = delayQueue(queueItem)(key)
                    'glf_debugLog.WriteToLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
                    GetRef(delayObject.Callback)(delayObject.Args)    
                End If
            Next
            delayQueue.Remove queueItem
        End If
    Next
End Sub
Function CreateGlfBallDevice(name)
	Dim device : Set device = (new GlfBallDevice)(name)
	Set CreateGlfBallDevice = device
End Function

Class GlfBallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeout
    Private m_balls
    Private m_balls_in_device
    Private m_eject_angle
    Private m_eject_angle_rnd
    Private m_eject_pitch
    Private m_eject_strength
    Private m_eject_strength_rnd
    Private m_eject_deltaz
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_balls_to_eject
    Private m_ejecting_all
    Private m_ejecting
    Private m_mechanical_eject
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
    
    Public Property Get Balls(): Balls = m_balls_in_device : End Property

    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    
    Public Property Let EjectAngle(value) : m_eject_angle = glf_PI * value / 180 : End Property
    Public Property Let EjectAngleRand(value) : m_eject_angle_rnd = glf_PI * value / 180 : End Property
    Public Property Let EjectPitch(value) : m_eject_pitch = glf_PI * value / 180 : End Property
    Public Property Let EjectStrength(value) : m_eject_strength = value : End Property
    Public Property Let EjectStrengthRand(value) : m_eject_strength_rnd = value : End Property
    Public Property Let EjectDeltaZ(value) : m_eject_deltaz = value : End Property
    
    Public Property Let EjectTimeout(value) : m_eject_timeout = value : End Property
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
    Public Property Let PlayerControlledEjectEvent(value)
        m_player_controlled_eject_event = value
        AddPinEventListener m_player_controlled_eject_event, m_name & "_eject_attempt", "BallDeviceEventHandler", 1000, Array("ball_eject", Me)
    End Property
    Public Property Let BallSwitches(value)
        m_ball_switches = value
        ReDim m_balls(Ubound(m_ball_switches))
        Dim x
        For x=0 to UBound(m_ball_switches)
            m_balls(x) = Null
            AddPinEventListener m_ball_switches(x)&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_entering", Me, x)
            AddPinEventListener m_ball_switches(x)&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, x)
        Next
    End Property
    Public Property Let MechanicalEject(value) : m_mechanical_eject = value : End Property
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
        m_eject_angle_rnd = 0
        m_eject_strength = 0
        m_eject_strength_rnd = 0
        m_eject_deltaz = 0
        m_ejecting = False
        m_eject_callback = Null
        m_ejecting_all = False
        m_balls_to_eject = 0
        m_balls_in_device = 0
        m_mechanical_eject = False
        m_eject_timeout = 1000
        glf_ball_devices.Add name, Me
	    Set Init = Me
	End Function

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_eject_timeout"
        'SoundSaucerLockAtBall ball
        Set m_balls(switch) = ball
        m_balls_in_device = m_balls_in_device + 1
        Log "Ball Entered" 
        Dim unclaimed_balls
        unclaimed_balls = DispatchRelayPinEvent(m_name & "_ball_enter", 1)
        Log "Unclaimed Balls: " & unclaimed_balls
        If (m_default_device = False Or m_ejecting = True) And unclaimed_balls > 0 And Not IsNull(m_balls(0)) Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        m_balls(switch) = Null
        m_balls_in_device = m_balls_in_device - 1
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechanical_eject = True And m_eject_timeout > 0 Then
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
                m_balls(0).VelX = (m_eject_strength + m_eject_strength_rnd*2*(rnd-0.5)) * Cos(m_eject_pitch) * Sin(m_eject_angle + m_eject_angle_rnd*2*(rnd-0.5))
                m_balls(0).VelY = (m_eject_strength + m_eject_strength_rnd*2*(rnd-0.5)) * Cos(m_eject_pitch) * Cos(m_eject_angle + m_eject_angle_rnd*2*(rnd-0.5)) * (-1)
                m_balls(0).VelZ = (m_eject_strength + m_eject_strength_rnd*2*(rnd-0.5)) * Sin(m_eject_pitch)
                m_balls(0).Z = m_balls(0).Z + m_eject_deltaz
                Log "VelX: " &  m_balls(0).VelX & ", VelY: " &  m_balls(0).VelY & ", VelZ: " &  m_balls(0).VelZ
            End If
        End If

        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(m_balls(0))
        End If
    End Sub

    Public Sub EjectBalls(balls)
        Log "Ejecting "&balls&" Balls."
        m_ejecting_all = True
        m_balls_to_eject = balls - 1
        Eject()
    End Sub

    Public Sub EjectAll
        Log "Ejecting All." 
        m_ejecting_all = True
        m_balls_to_eject = m_balls_in_device - 1 
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
Function CreateGlfDiverter(name)
	Dim diverter : Set diverter = (new GlfDiverter)(name)
	Set CreateGlfDiverter = diverter
End Function

Class GlfDiverter

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
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "DiverterEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "DiverterEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let ActivateEvents(value) : m_activate_events = value : End Property
    Public Property Let DeactivateEvents(value) : m_deactivate_events = value : End Property
    Public Property Let ActivationTime(value) : m_activation_time = value : End Property
    Public Property Let ActivationSwitches(value) : m_activation_switches = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "diverter_" & name
        m_enable_events = Array()
        m_disable_events = Array()
        m_activate_events = Array()
        m_deactivate_events = Array()
        m_activation_switches = Array()
        m_activation_time = 0
        m_debug = False
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
Function CreateGlfDroptarget(name)
	Dim droptarget : Set droptarget = (new GlfDroptarget)(name)
	Set CreateGlfDroptarget = droptarget
End Function

Class GlfDroptarget

    Private m_name
	Private m_switch
    Private m_enable_keep_up_events
    Private m_disable_keep_up_events
	Private m_action_cb
	Private m_knockdown_events
	Private m_reset_events

    
    Private m_debug

	Public Property Let Switch(value)
		m_switch = value
		AddPinEventListener m_switch & "_active", m_name & "_switch_active", "DroptargetEventHandler", 1000, Array("switch_active", Me)
		AddPinEventListener m_switch & "_inactive", m_name & "_switch_inactive", "DroptargetEventHandler", 1000, Array("switch_inactive", Me)
	End Property
    Public Property Let EnableKeepUpEvents(value)
        Dim evt
        If IsArray(m_enable_keep_up_events) Then
            For Each evt in m_enable_keep_up_events
                RemovePinEventListener evt, m_name & "_enable_keepup"
            Next
        End If
        m_enable_keep_up_events = value
        For Each evt in m_enable_keep_up_events
            AddPinEventListener evt, m_name & "_enable_keepup", "DroptargetEventHandler", 1000, Array("enable_keepup", Me)
        Next
    End Property
    Public Property Let DisableKeepUpEvents(value)
        Dim evt
        If IsArray(m_disable_keep_up_events) Then
            For Each evt in m_disable_keep_up_events
                RemovePinEventListener evt, m_name & "_disable_keepup"
            Next
        End If
        m_disable_keep_up_events = value
        For Each evt in m_disable_keep_up_events
            AddPinEventListener evt, m_name & "_disable_keepup", "DroptargetEventHandler", 1000, Array("disable_keepup", Me)
        Next
    End Property

    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
	Public Property Let KnockdownEvents(value)
		Dim evt
		If IsArray(m_knockdown_events) Then
			For Each evt in m_knockdown_events
				RemovePinEventListener evt, m_name & "_knockdown"
			Next
		End If
		m_knockdown_events = value
		For Each evt in m_knockdown_events
			AddPinEventListener evt, m_name & "_knockdown", "DroptargetEventHandler", 1000, Array("knockdown", Me)
		Next
	End Property
	Public Property Let ResetEvents(value)
		Dim evt
		If IsArray(m_reset_events) Then
			For Each evt in m_reset_events
				RemovePinEventListener evt, m_name & "_reset"
			Next
		End If
		m_reset_events = value
		For Each evt in m_reset_events
			AddPinEventListener evt, m_name & "_reset", "DroptargetEventHandler", 1000, Array("reset", Me)
		Next
	End Property

	Public default Function init(name)
        m_name = "drop_target_" & name
		m_switch = Empty
        EnableKeepUpEventsEnableEvents = Array()
        DisableKeepUpEventsEnableEvents = Array()
		m_action_cb = Empty
		KnockdownEventsEvents = Array()
		ResetEventsEvents = Array()


		
		
		m_debug = False
        glf_droptargets.Add name, Me
        Set Init = Me
	End Function

    Public Sub UpdateStateFromSwitch(is_complete)

		Log "Drop target " & m_name & " switch " & m_switch & " has active value " & isComplete & " compared to drop complete " & m_complete

		If is_complete <> m_complete Then
			If is_complete = True Then
				Down()
			Else
				Up()
			End	If
		End If
		UpdateBanks()
    End Sub

    Public Sub Up()
        m_complete = False
        DispatchPinEvent name & "_up", Null
    End Sub

	Public Sub Down()
        m_complete = True
        DispatchPinEvent name & "_up", Null
    End Sub

	Public Sub EnableKeepup()
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(3)
		End If
    End Sub

	Public Sub DisableKeepup()
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(4)
		End If
    End Sub

	Public Sub Knockdown()
        If Not IsEmpty(m_action_cb) And m_complete = False Then
            GetRef(m_action_cb)(1)
		End If
    End Sub

	Public Sub Reset()
        If Not IsEmpty(m_action_cb) And m_complete = True Then
            GetRef(m_action_cb)(0)
		End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DroptargetEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim droptarget : Set droptarget = ownProps(1)
    Select Case evt
        Case "switch_active"
            droptarget.UpdateStateFromSwitch 1
        Case "switch_inactive"
            droptarget.UpdateStateFromSwitch 0
        Case "enable_keepup"
            droptarget.EnableKeepup
        Case "disable_keepup"
            droptarget.DisableKeepup
        Case "knockdown"
            droptarget.Knockdown
        Case "reset"
            droptarget.Reset
    End Select
    If IsObject(args(1)) Then
        Set DroptargetEventHandler = kwargs
    Else
        DroptargetEventHandler = kwargs
    End If
    
End Function

Function CreateGlfFlipper(name)
	Dim flipper : Set flipper = (new GlfFlipper)(name)
	Set CreateGlfFlipper = flipper
End Function

Class GlfFlipper

    Private m_name
    Private m_enable_events
    Private m_disable_events
    Private m_enabled
    Private m_switches
    Private m_action_cb
    Private m_debug

    Public Property Let Switch(value)
        m_switches = value
    End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "FlipperEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "FlipperEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "flipper_" & name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_enabled = False
        m_action_cb = Empty
        m_switches = Array()
        m_debug = False
        glf_flippers.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_switches
            AddPinEventListener evt & "_active", m_name & "_active", "FlipperEventHandler", 1000, Array("activate", Me)
            AddPinEventListener evt & "_inactive", m_name & "_inactive", "FlipperEventHandler", 1000, Array("deactivate", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Deactivate()
        Dim evt
        For Each evt in m_switches
            RemovePinEventListener evt & "_active", m_name & "_active"
            RemovePinEventListener evt & "_inactive", m_name & "_inactive"
        Next
    End Sub

    Public Sub Activate()
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(1)
        End If
        DispatchPinEvent m_name & "_activate", Null
    End Sub

    Public Sub Deactivate()
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(0)
        End If
        DispatchPinEvent m_name & "_deactivate", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function FlipperEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim flipper : Set flipper = ownProps(1)
    Select Case evt
        Case "enable"
            flipper.Enable
        Case "disable"
            flipper.Disable
        Case "activate"
            flipper.Activate
        Case "deactivate"
            flipper.Deactivate
    End Select
    FlipperEventHandler = kwargs
End Function
Function CreateGlfLightSegmentDisplay(name)
	Dim segment_display : Set segment_display = (new GlfLightSegmentDisplay)(name)
	Set CreateGlfLightSegmentDisplay = segment_display
End Function

Class GlfLightSegmentDisplay

    private m_flash_on
    private m_flashing
    private m_flash_mask
    private m_text
    private m_current_text
    private m_display_state
    private m_lights
    private m_light_group
    private m_segmentmap
    private m_segment_type
    private m_size

    Public Property Get SegmentType() : SegmentType = m_segment_type : End Property
    Public Property Let SegmentType(input)
        m_segment_type = input
        If m_segment_type = "14Segment" Then
            Set m_segmentmap = FOURTEEN_SEGMENTS
        ElseIf m_segment_type = "7Segment" Then
            Set m_segmentmap = SEVEN_SEGMENTS
        End If
        CalculateLights()
    End Property

    Public Property Get LightGroup() : LightGroup = m_light_group : End Property
    Public Property Let LightGroup(input)
        m_light_group = input
        CalculateLights()
    End Property

    Public Property Get SegmentSize() : SegmentSize = m_size : End Property
    Public Property Let SegmentSize(input)
        m_size = input
        CalculateLights()
    End Property

    Public default Function init(name)

        m_flash_on = True
        m_flashing = "no_flash"
        m_flash_mask = Empty
        m_text = Empty
        m_size = 0
        m_segment_type = Empty
        m_segmentmap = Null
        m_light_group = Empty
        m_current_text = Empty
        m_display_state = Empty
        m_lights = Array()  

        glf_segment_displays.Add name, Me
        Set Init = Me
    End Function

    Private Sub CalculateLights()
        If Not IsEmpty(m_segment_type) And m_size > 0 And Not IsEmpty(m_light_group) Then
            m_lights = Array()
            If m_segment_type = "14Segment" Then
                ReDim m_lights(m_size * 15)
            ElseIf m_segment_type = "7Segment" Then
                ReDim m_lights(m_size * 8)
            End If

            Dim i
            For i=0 to UBound(m_lights)
                m_lights(i) = m_light_group & CStr(i+1)
            Next
        End If
    End Sub

    Private Sub SetText(text, flashing, flash_mask)
        'Set a text to the display.
        m_text = text
        m_flashing = flashing

        If flashing = "no_flash" Then
            m_flash_on = True
        ElseIf flashing = "flash_mask" Then
            'm_flash_mask = flash_mask.rjust(len(text))
        End If

        If flashing = "no_flash" or m_flash_on = True or Not IsEmpty(text) Then
            If text <> m_display_state Then
                m_display_state = text
                'Set text to lights.
                If text="" Then
                    text = Glf_FormatValue(text, " >" & CStr(m_size))
                Else
                    text = Right(text, m_size)
                End If
                If text <> m_current_text Then
                    m_current_text = text
                    UpdateText()
                End If
            End If
        End If
    End Sub

    Private Sub UpdateText()
        'iterate lights and chars
        Dim mapped_text, segment
        mapped_text = MapSegmentTextToSegments(m_current_text, m_size, m_segmentmap)
        Dim segment_idx : segment_idx = 1
        For Each segment in mapped_text
            
            If m_segment_type = "14Segment" Then
                Glf_SetLight m_light_group & CStr(segment_idx), SegmentColor(segment.a)
                Glf_SetLight m_light_group & CStr(segment_idx + 1), SegmentColor(segment.b)
                Glf_SetLight m_light_group & CStr(segment_idx + 2), SegmentColor(segment.c)
                Glf_SetLight m_light_group & CStr(segment_idx + 3), SegmentColor(segment.d)
                Glf_SetLight m_light_group & CStr(segment_idx + 4), SegmentColor(segment.e)
                Glf_SetLight m_light_group & CStr(segment_idx + 5), SegmentColor(segment.f)
                Glf_SetLight m_light_group & CStr(segment_idx + 6), SegmentColor(segment.g1)
                Glf_SetLight m_light_group & CStr(segment_idx + 7), SegmentColor(segment.g2)
                Glf_SetLight m_light_group & CStr(segment_idx + 8), SegmentColor(segment.h)
                Glf_SetLight m_light_group & CStr(segment_idx + 9), SegmentColor(segment.j)
                Glf_SetLight m_light_group & CStr(segment_idx + 10), SegmentColor(segment.k)
                Glf_SetLight m_light_group & CStr(segment_idx + 11), SegmentColor(segment.n)
                Glf_SetLight m_light_group & CStr(segment_idx + 12), SegmentColor(segment.m)
                Glf_SetLight m_light_group & CStr(segment_idx + 13), SegmentColor(segment.l)
                Glf_SetLight m_light_group & CStr(segment_idx + 14), SegmentColor(segment.dp)

            ElseIf m_segment_type = "7Segment" Then
                
            End If
            
            segment_idx = segment_idx + 15
        Next
        'for char, lights_for_char in zip(mapped_text, self._lights):
        '    for name, light in lights_for_char.items():
        '        if getattr(char[0], name):
        '            light.color(color=char[1], key=self._key)
        '        else:
        '            light.remove_from_stack_by_key(key=self._key)
    End Sub

    Private Function SegmentColor(value)
        If value = 1 Then
            SegmentColor = "ffffff"
        Else
            SegmentColor = "000000"
        End If
    End Function

    Public Sub AddTextEntry(text, color, flashing, flash_mask, transition, transition_out, priority, key)
        SetText text, "no_flash", Empty
    End Sub

    Public Sub RemoveTextByKey(key)

    End Sub

    Public Sub SetFlashing(flash_type)

    End Sub

    Public Sub SetFlashingMask(mask)

    End Sub

    Public Sub SetColor(color)

    End Sub

End Class

Class FourteenSegments
    Public dp, l, m, n, k, j, h, g2, g1, f, e, d, c, b, a, char

    Public default Function init(dp, l, m, n, k, j, h, g2, g1, f, e, d, c, b, a, char)
        Me.dp = dp
        Me.a = a
        Me.b = b
        Me.c = c
        Me.d = d
        Me.e = e
        Me.f = f
        Me.g1 = g1
        Me.g2 = g2
        Me.h = h
        Me.j = j
        Me.k = k
        Me.n = n
        Me.m = m
        Me.l = l
        Me.char = char
        Set Init = Me
    End Function
End Class

Class SevenSegments
    Public dp, g, f, e, d, c, b, a, char

    Public default Function init(dp, g, f, e, d, c, b, a, char)
        Me.dp = dp
        Me.a = a
        Me.b = b
        Me.c = c
        Me.d = d
        Me.e = e
        Me.f = f
        Me.g = g
        Me.char = char
        Set Init = Me
    End Function
End Class


Dim FOURTEEN_SEGMENTS
Set FOURTEEN_SEGMENTS = CreateObject("Scripting.Dictionary")

FOURTEEN_SEGMENTS.Add Null, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "?")
FOURTEEN_SEGMENTS.Add 32, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " ")
FOURTEEN_SEGMENTS.Add 33, (New FourteenSegments)(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, "!")
FOURTEEN_SEGMENTS.Add 34, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, Chr(34)) ' Character "
FOURTEEN_SEGMENTS.Add 35, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 0, 1, 1, 1, 0, "#")
FOURTEEN_SEGMENTS.Add 36, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 1, 1, 0, 1, 1, 0, 1, "$")
FOURTEEN_SEGMENTS.Add 37, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 1, 1, 1, 0, 0, 1, 0, 0, "%")
FOURTEEN_SEGMENTS.Add 38, (New FourteenSegments)(0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 1, "&")
FOURTEEN_SEGMENTS.Add 39, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "'")
FOURTEEN_SEGMENTS.Add 40, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "(")
FOURTEEN_SEGMENTS.Add 41, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, ")")
FOURTEEN_SEGMENTS.Add 42, (New FourteenSegments)(0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, "*")
FOURTEEN_SEGMENTS.Add 43, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, "+")
FOURTEEN_SEGMENTS.Add 44, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ",")
FOURTEEN_SEGMENTS.Add 45, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, "-")
FOURTEEN_SEGMENTS.Add 46, (New FourteenSegments)(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ".")
FOURTEEN_SEGMENTS.Add 47, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "/")
FOURTEEN_SEGMENTS.Add 48, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "0")
FOURTEEN_SEGMENTS.Add 49, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, "1")
FOURTEEN_SEGMENTS.Add 50, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 1, "2")
FOURTEEN_SEGMENTS.Add 51, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 1, "3")
FOURTEEN_SEGMENTS.Add 52, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, 1, 0, "4")
FOURTEEN_SEGMENTS.Add 53, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 0, 1, "5")
FOURTEEN_SEGMENTS.Add 54, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 1, "6")
FOURTEEN_SEGMENTS.Add 55, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, "7")
FOURTEEN_SEGMENTS.Add 56, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, "8")
FOURTEEN_SEGMENTS.Add 57, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1, 1, "9")
FOURTEEN_SEGMENTS.Add 58, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, ":")
FOURTEEN_SEGMENTS.Add 59, (New FourteenSegments)(0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, ";")
FOURTEEN_SEGMENTS.Add 60, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, "<")
FOURTEEN_SEGMENTS.Add 61, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, "=")
FOURTEEN_SEGMENTS.Add 62, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, ">")
FOURTEEN_SEGMENTS.Add 63, (New FourteenSegments)(1, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, "?")
FOURTEEN_SEGMENTS.Add 64, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 1, "@")
FOURTEEN_SEGMENTS.Add 65, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, "A")
FOURTEEN_SEGMENTS.Add 66, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 1, 1, "B")
FOURTEEN_SEGMENTS.Add 67, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, "C")
FOURTEEN_SEGMENTS.Add 68, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, "D")
FOURTEEN_SEGMENTS.Add 69, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, "E")
FOURTEEN_SEGMENTS.Add 70, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 1, "F")
FOURTEEN_SEGMENTS.Add 71, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 0, 1, "G")
FOURTEEN_SEGMENTS.Add 72, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 0, "H")
FOURTEEN_SEGMENTS.Add 73, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 1, "I")
FOURTEEN_SEGMENTS.Add 74, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, "J")
FOURTEEN_SEGMENTS.Add 75, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, "K")
FOURTEEN_SEGMENTS.Add 76, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, "L")
FOURTEEN_SEGMENTS.Add 77, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, "M")
FOURTEEN_SEGMENTS.Add 78, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, "N")
FOURTEEN_SEGMENTS.Add 79, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "O")
FOURTEEN_SEGMENTS.Add 80, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 1, "P")
FOURTEEN_SEGMENTS.Add 81, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "Q")
FOURTEEN_SEGMENTS.Add 82, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 1, "R")
FOURTEEN_SEGMENTS.Add 83, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 0, 1, "S")
FOURTEEN_SEGMENTS.Add 84, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, "T")
FOURTEEN_SEGMENTS.Add 85, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, "U")
FOURTEEN_SEGMENTS.Add 86, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, "V")
FOURTEEN_SEGMENTS.Add 87, (New FourteenSegments)(0, 1, 0, 1, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, "W")
FOURTEEN_SEGMENTS.Add 88, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "X")
FOURTEEN_SEGMENTS.Add 89, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1, 1, 1, 0, "Y")
FOURTEEN_SEGMENTS.Add 90, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, "Z")
FOURTEEN_SEGMENTS.Add 91, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, "[")
FOURTEEN_SEGMENTS.Add 92, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, Chr(92)) ' Character \
FOURTEEN_SEGMENTS.Add 93, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, "]")
FOURTEEN_SEGMENTS.Add 94, (New FourteenSegments)(0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "^")
FOURTEEN_SEGMENTS.Add 95, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, "_")
FOURTEEN_SEGMENTS.Add 96, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "`")
FOURTEEN_SEGMENTS.Add 97, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 1, 0, 0, 0, "a")
FOURTEEN_SEGMENTS.Add 98, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, "b")
FOURTEEN_SEGMENTS.Add 99, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 0, 0, "c")
FOURTEEN_SEGMENTS.Add 100, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, "d")
FOURTEEN_SEGMENTS.Add 101, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 1, 1, 0, 0, 0, "e")
FOURTEEN_SEGMENTS.Add 102, (New FourteenSegments)(0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, "f")
FOURTEEN_SEGMENTS.Add 103, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 1, 1, 1, 1, "g")
FOURTEEN_SEGMENTS.Add 104, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 0, 0, "h")
FOURTEEN_SEGMENTS.Add 105, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "i")
FOURTEEN_SEGMENTS.Add 106, (New FourteenSegments)(0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, "j")
FOURTEEN_SEGMENTS.Add 107, (New FourteenSegments)(0, 1, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "k")
FOURTEEN_SEGMENTS.Add 108, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "l")
FOURTEEN_SEGMENTS.Add 109, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 0, 0, "m")
FOURTEEN_SEGMENTS.Add 110, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 0, 0, "n")
FOURTEEN_SEGMENTS.Add 111, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 0, "o")
FOURTEEN_SEGMENTS.Add 112, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 0, 1, 1, "p")
FOURTEEN_SEGMENTS.Add 113, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 1, 1, 1, 1, "q")
FOURTEEN_SEGMENTS.Add 114, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 0, 0, 0, "r")
FOURTEEN_SEGMENTS.Add 115, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, "s")
FOURTEEN_SEGMENTS.Add 116, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, "t")
FOURTEEN_SEGMENTS.Add 117, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, "u")
FOURTEEN_SEGMENTS.Add 118, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, "v")
FOURTEEN_SEGMENTS.Add 119, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, 0, "w")
FOURTEEN_SEGMENTS.Add 120, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "x")
FOURTEEN_SEGMENTS.Add 121, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 1, 1, 0, "y")
FOURTEEN_SEGMENTS.Add 122, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 1, "z")
FOURTEEN_SEGMENTS.Add 123, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 1, 0, 1, 0, 0, 1, 0, 0, 1, "{")
FOURTEEN_SEGMENTS.Add 124, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "|")
FOURTEEN_SEGMENTS.Add 125, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1, 1, 0, 0, 1, "}")
FOURTEEN_SEGMENTS.Add 126, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, "~")


Dim SEVEN_SEGMENTS
Set SEVEN_SEGMENTS = CreateObject("Scripting.Dictionary")

SEVEN_SEGMENTS.Add Null, (New SevenSegments)(0, 0, 0, 0, 0, 0, 0, 0, "?")
SEVEN_SEGMENTS.Add 32, (New SevenSegments)(0, 0, 0, 0, 0, 0, 0, 0, " ")
SEVEN_SEGMENTS.Add 33, (New SevenSegments)(1, 0, 0, 0, 0, 1, 1, 0, "!")
SEVEN_SEGMENTS.Add 34, (New SevenSegments)(0, 0, 1, 0, 0, 0, 1, 0, Chr(34)) ' Character "
SEVEN_SEGMENTS.Add 35, (New SevenSegments)(0, 1, 1, 1, 1, 1, 1, 0, "#")
SEVEN_SEGMENTS.Add 36, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "$")
SEVEN_SEGMENTS.Add 37, (New SevenSegments)(1, 1, 0, 1, 0, 0, 1, 0, "%")
SEVEN_SEGMENTS.Add 38, (New SevenSegments)(0, 1, 0, 0, 0, 1, 1, 0, "&")
SEVEN_SEGMENTS.Add 39, (New SevenSegments)(0, 0, 1, 0, 0, 0, 0, 0, "'")
SEVEN_SEGMENTS.Add 40, (New SevenSegments)(0, 0, 1, 0, 1, 0, 0, 1, "(")
SEVEN_SEGMENTS.Add 41, (New SevenSegments)(0, 0, 0, 0, 1, 0, 1, 1, ")")
SEVEN_SEGMENTS.Add 42, (New SevenSegments)(0, 0, 1, 0, 0, 0, 0, 1, "*")
SEVEN_SEGMENTS.Add 43, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 0, "+")
SEVEN_SEGMENTS.Add 44, (New SevenSegments)(0, 0, 0, 1, 0, 0, 0, 0, ",")
SEVEN_SEGMENTS.Add 45, (New SevenSegments)(0, 1, 0, 0, 0, 0, 0, 0, "-")
SEVEN_SEGMENTS.Add 46, (New SevenSegments)(1, 0, 0, 0, 0, 0, 0, 0, ".")
SEVEN_SEGMENTS.Add 47, (New SevenSegments)(0, 1, 0, 1, 0, 0, 1, 0, "/")
SEVEN_SEGMENTS.Add 48, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 1, "0")
SEVEN_SEGMENTS.Add 49, (New SevenSegments)(0, 0, 0, 0, 0, 1, 1, 0, "1")
SEVEN_SEGMENTS.Add 50, (New SevenSegments)(0, 1, 0, 1, 1, 0, 1, 1, "2")
SEVEN_SEGMENTS.Add 51, (New SevenSegments)(0, 1, 0, 0, 1, 1, 1, 1, "3")
SEVEN_SEGMENTS.Add 52, (New SevenSegments)(0, 1, 1, 0, 0, 1, 1, 0, "4")
SEVEN_SEGMENTS.Add 53, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "5")
SEVEN_SEGMENTS.Add 54, (New SevenSegments)(0, 1, 1, 1, 1, 1, 0, 1, "6")
SEVEN_SEGMENTS.Add 55, (New SevenSegments)(0, 0, 0, 0, 0, 1, 1, 1, "7")
SEVEN_SEGMENTS.Add 56, (New SevenSegments)(0, 1, 1, 1, 1, 1, 1, 1, "8")
SEVEN_SEGMENTS.Add 57, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 1, "9")
SEVEN_SEGMENTS.Add 58, (New SevenSegments)(0, 0, 0, 0, 1, 0, 0, 1, ":")
SEVEN_SEGMENTS.Add 59, (New SevenSegments)(0, 0, 0, 0, 1, 1, 0, 1, ";")
SEVEN_SEGMENTS.Add 60, (New SevenSegments)(0, 1, 1, 0, 0, 0, 0, 1, "<")
SEVEN_SEGMENTS.Add 61, (New SevenSegments)(0, 1, 0, 0, 1, 0, 0, 0, "=")
SEVEN_SEGMENTS.Add 62, (New SevenSegments)(0, 1, 0, 0, 0, 0, 1, 1, ">")
SEVEN_SEGMENTS.Add 63, (New SevenSegments)(1, 1, 0, 1, 0, 0, 1, 1, "?")
SEVEN_SEGMENTS.Add 64, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 1, "@")
SEVEN_SEGMENTS.Add 65, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 1, "A")
SEVEN_SEGMENTS.Add 66, (New SevenSegments)(0, 1, 1, 1, 1, 1, 0, 0, "B")
SEVEN_SEGMENTS.Add 67, (New SevenSegments)(0, 0, 1, 1, 1, 0, 0, 1, "C")
SEVEN_SEGMENTS.Add 68, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 0, "D")
SEVEN_SEGMENTS.Add 69, (New SevenSegments)(0, 1, 1, 1, 1, 0, 0, 1, "E")
SEVEN_SEGMENTS.Add 70, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 1, "F")
SEVEN_SEGMENTS.Add 71, (New SevenSegments)(0, 0, 1, 1, 1, 1, 0, 1, "G")
SEVEN_SEGMENTS.Add 72, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 0, "H")
SEVEN_SEGMENTS.Add 73, (New SevenSegments)(0, 0, 1, 1, 0, 0, 0, 0, "I")
SEVEN_SEGMENTS.Add 74, (New SevenSegments)(0, 0, 0, 1, 1, 1, 1, 0, "J")
SEVEN_SEGMENTS.Add 75, (New SevenSegments)(0, 1, 1, 1, 0, 1, 0, 1, "K")
SEVEN_SEGMENTS.Add 76, (New SevenSegments)(0, 0, 1, 1, 1, 0, 0, 0, "L")
SEVEN_SEGMENTS.Add 77, (New SevenSegments)(0, 0, 0, 1, 0, 1, 0, 1, "M")
SEVEN_SEGMENTS.Add 78, (New SevenSegments)(0, 0, 1, 1, 0, 1, 1, 1, "N")
SEVEN_SEGMENTS.Add 79, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 1, "O")
SEVEN_SEGMENTS.Add 80, (New SevenSegments)(0, 1, 1, 1, 0, 0, 1, 1, "P")
SEVEN_SEGMENTS.Add 81, (New SevenSegments)(0, 1, 1, 0, 1, 0, 1, 1, "Q")
SEVEN_SEGMENTS.Add 82, (New SevenSegments)(0, 0, 1, 1, 0, 0, 1, 1, "R")
SEVEN_SEGMENTS.Add 83, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "S")
SEVEN_SEGMENTS.Add 84, (New SevenSegments)(0, 1, 1, 1, 1, 0, 0, 0, "T")
SEVEN_SEGMENTS.Add 85, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 0, "U")
SEVEN_SEGMENTS.Add 86, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 0, "V")
SEVEN_SEGMENTS.Add 87, (New SevenSegments)(0, 0, 1, 0, 1, 0, 1, 0, "W")
SEVEN_SEGMENTS.Add 88, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 0, "X")
SEVEN_SEGMENTS.Add 89, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 0, "Y")
SEVEN_SEGMENTS.Add 90, (New SevenSegments)(0, 1, 0, 1, 1, 0, 1, 1, "Z")
SEVEN_SEGMENTS.Add 91, (New SevenSegments)(0, 0, 1, 1, 1, 0, 0, 1, "[")
SEVEN_SEGMENTS.Add 92, (New SevenSegments)(0, 1, 1, 0, 0, 1, 0, 0, Chr(92)) ' Character \
SEVEN_SEGMENTS.Add 93, (New SevenSegments)(0, 0, 0, 0, 1, 1, 1, 1, "]")
SEVEN_SEGMENTS.Add 94, (New SevenSegments)(0, 0, 1, 0, 0, 0, 1, 1, "^")
SEVEN_SEGMENTS.Add 95, (New SevenSegments)(0, 0, 0, 0, 1, 0, 0, 0, "_")
SEVEN_SEGMENTS.Add 96, (New SevenSegments)(0, 0, 0, 0, 0, 0, 1, 0, "`")
SEVEN_SEGMENTS.Add 97, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 1, "a")
SEVEN_SEGMENTS.Add 98, (New SevenSegments)(0, 1, 1, 1, 1, 1, 0, 0, "b")
SEVEN_SEGMENTS.Add 99, (New SevenSegments)(0, 1, 0, 1, 1, 0, 0, 0, "c")
SEVEN_SEGMENTS.Add 100, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 0, "d")
SEVEN_SEGMENTS.Add 101, (New SevenSegments)(0, 1, 1, 1, 1, 0, 1, 1, "e")
SEVEN_SEGMENTS.Add 102, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 1, "f")
SEVEN_SEGMENTS.Add 103, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 1, "g")
SEVEN_SEGMENTS.Add 104, (New SevenSegments)(0, 1, 1, 1, 0, 1, 0, 0, "h")
SEVEN_SEGMENTS.Add 105, (New SevenSegments)(0, 0, 0, 1, 0, 0, 0, 0, "i")
SEVEN_SEGMENTS.Add 106, (New SevenSegments)(0, 0, 0, 0, 1, 1, 0, 0, "j")
SEVEN_SEGMENTS.Add 107, (New SevenSegments)(0, 1, 1, 1, 0, 1, 0, 1, "k")
SEVEN_SEGMENTS.Add 108, (New SevenSegments)(0, 0, 1, 1, 0, 0, 0, 0, "l")
SEVEN_SEGMENTS.Add 109, (New SevenSegments)(0, 0, 0, 1, 0, 1, 0, 0, "m")
SEVEN_SEGMENTS.Add 110, (New SevenSegments)(0, 1, 0, 1, 0, 1, 0, 0, "n")
SEVEN_SEGMENTS.Add 111, (New SevenSegments)(0, 1, 0, 1, 1, 1, 0, 0, "o")
SEVEN_SEGMENTS.Add 112, (New SevenSegments)(0, 1, 1, 1, 0, 0, 1, 1, "p")
SEVEN_SEGMENTS.Add 113, (New SevenSegments)(0, 1, 1, 0, 0, 1, 1, 1, "q")
SEVEN_SEGMENTS.Add 114, (New SevenSegments)(0, 1, 0, 1, 0, 0, 0, 0, "r")
SEVEN_SEGMENTS.Add 115, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "s")
SEVEN_SEGMENTS.Add 116, (New SevenSegments)(0, 1, 1, 1, 1, 0, 0, 0, "t")
SEVEN_SEGMENTS.Add 117, (New SevenSegments)(0, 0, 0, 1, 1, 1, 0, 0, "u")
SEVEN_SEGMENTS.Add 118, (New SevenSegments)(0, 0, 0, 1, 1, 1, 0, 0, "v")
SEVEN_SEGMENTS.Add 119, (New SevenSegments)(0, 0, 0, 1, 0, 1, 0, 0, "w")
SEVEN_SEGMENTS.Add 120, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 0, "x")
SEVEN_SEGMENTS.Add 121, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 0, "y")
SEVEN_SEGMENTS.Add 122, (New SevenSegments)(0, 1, 0, 1, 1, 0, 1, 1, "z")
SEVEN_SEGMENTS.Add 123, (New SevenSegments)(0, 1, 0, 0, 0, 1, 1, 0, "{")
SEVEN_SEGMENTS.Add 124, (New SevenSegments)(0, 0, 1, 1, 0, 0, 0, 0, "|")
SEVEN_SEGMENTS.Add 125, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 0, "}")
SEVEN_SEGMENTS.Add 126, (New SevenSegments)(0, 0, 0, 0, 0, 0, 0, 1, "~")


Function MapSegmentTextToSegments(text, display_width, segment_mapping)
    'Map a segment display text to a certain display mapping.
    Dim segments()
	ReDim segments(Len(text)-1)

	Dim charCode, i
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        charCode = Asc(char)
        If segment_mapping.Exists(CharCode) Then
            Set mapping = segment_mapping(CharCode)
        Else
            Set mapping = segment_mapping(Null)
        End If
        Set segments(i-1) = mapping
    Next

    MapSegmentTextToSegments = segments
End Function
Function CreateGlfMagnet(name)
	Dim magnet : Set magnet = (new GlfMagnet)(name)
	Set CreateGlfMagnet = magnet
End Function

Class GlfMagnet

    Private m_name
    Private m_enable_events
    Private m_disable_events
    Private m_fling_ball_events
    Private m_fling_drop_time
    Private m_fling_regrab_time
    Private m_grab_ball_events
    Private m_grab_switch
    Private m_grab_time
    Private m_release_ball_events
    Private m_release_time
    Private m_reset_events
    Private m_action_cb

    Private m_active
    Private m_release_in_progress

    Private m_debug

    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MagnetEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "MagnetEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let FlingBallEvents(value) : m_fling_ball_events = value : End Property
    Public Property Let FlingDropTime(value) : Set m_fling_drop_time = CreateGlfInput(value) : End Property
    Public Property Let FlingRegrabTime(value) : Set m_fling_regrab_time = CreateGlfInput(value) : End Property
    Public Property Let GrabBallEvents(value) : m_grab_ball_events = value : End Property
    Public Property Let GrabSwitch(value)
        m_grab_switch = value
    End Property
    Public Property Let GrabTime(value) : Set m_grab_time = CreateGlfInput(value) : End Property
    Public Property Let ReleaseBallEvents(value) : m_release_ball_events = value : End Property
    Public Property Let ReleaseTime(value) : Set m_release_time = CreateGlfInput(value) : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "magnet_" & name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_action_cb = Empty
        m_fling_ball_events = Array()
        Set m_fling_drop_time = CreateGlfInput(250)
        Set m_fling_regrab_time = CreateGlfInput(50)
        m_grab_ball_events = Array()
        m_grab_switch = Empty
        Set m_grab_time = CreateGlfInput(1500)
        m_release_ball_events = Array()
        Set m_release_time = CreateGlfInput(500)
        m_reset_events = Array("machine_reset_phase_3", "ball_starting")
        m_active = False
        m_release_in_progress = False
        m_debug = False
        glf_magnets.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_fling_ball_events
            AddPinEventListener evt, m_name & "_fling", "MagnetEventHandler", 1000, Array("fling", Me)
        Next
        For Each evt in m_grab_ball_events
            AddPinEventListener evt, m_name & "_grab", "MagnetEventHandler", 1000, Array("grab", Me)
        Next
        For Each evt in m_release_ball_events
            AddPinEventListener evt, m_name & "_release", "MagnetEventHandler", 1000, Array("release", Me)
        Next
        AddPinEventListener m_grab_switch & "_active", m_name & "_grab_switch", "MagnetEventHandler", 1000, Array("grab", Me)
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_fling_ball_events
            RemovePinEventListener evt, m_name & "_fling"
        Next
        For Each evt in m_grab_ball_events
            RemovePinEventListener evt, m_name & "_grab"
        Next
        For Each evt in m_release_ball_events
            RemovePinEventListener evt, m_name & "_release"
        Next
        RemovePinEventListener m_grab_switch & "_active", m_name & "_grab_switch"
    End Sub

    Public Sub AddBall(ball)
        m_magnet.AddBall ball
    End Sub

    Public Sub RemoveBall(ball)
        m_magnet.RemoveBall ball
    End Sub

    Public Sub Fling()
        If m_active = False or m_release_in_progress = True Then
            Exit Sub
        End If
        m_active = False
        m_release_in_progress = True
        Log "Flinging Ball"
        DispatchPinEvent m_name & "flinging_ball", Null
        GetRef(m_action_cb)(0)
        SetDelay m_name & "_fling_reenable", "MagnetEventHandler" , Array(Array("fling_reenable", Me),Null), m_fling_drop_time.Value
    End Sub

    Public Sub FlingReenable()
        GetRef(m_action_cb)(1)
        SetDelay m_name & "_fling_regrab", "MagnetEventHandler" , Array(Array("fling_regrab", Me),Null), m_fling_regrab_time.Value
    End Sub

    Public Sub FlingRegrab()
        m_release_in_progress = False
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_flinged_ball", Null
    End Sub

    Public Sub Grab()
        If m_active = True Or m_release_in_progress = True Then
            Exit Sub
        End If
        Log "Grabbing Ball"
        m_active = True
        GetRef(m_action_cb)(1)
        DispatchPinEvent m_name & "_grabbing_ball", Null
        SetDelay m_name & "_grabbing_done", "MagnetEventHandler" , Array(Array("grabbing_done", Me),Null), m_grab_time.Value
    End Sub

    Public Sub GrabbingDone()
        DispatchPinEvent m_name & "_grabbed_ball", Null
    End Sub

    Public Sub Release()
        If m_active = False or m_release_in_progress = True Then
            Exit Sub
        End If
        m_active = False
        m_release_in_progress = True
        Log "Releasing Ball"
        DispatchPinEvent m_name & "releasing_ball", Null
        GetRef(m_action_cb)(0)
        SetDelay m_name & "_release_done", "MagnetEventHandler" , Array(Array("release_done", Me),Null), m_release_time.Value
    End Sub

    Public Sub ReleaseDone()
        m_release_in_progress = False
        DispatchPinEvent m_name & "released_ball", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MagnetEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim magnet : Set magnet = ownProps(1)
    Select Case evt
        Case "enable"
            magnet.Enable
        Case "disable"
            magnet.Disable
        Case "fling"
            magnet.Fling
        Case "grab"
            magnet.Grab
        Case "release"
            magnet.Release
        Case "grabbing_done"
            magnet.GrabbingDone
        Case "release_done"
            magnet.ReleaseDone
        Case "fling_reenable"
            magnet.FlingReenable
        Case "fling_regrab"
            magnet.FlingRegrab
    End Select
    If IsObject(args(1)) Then
        Set MagnetEventHandler = kwargs
    Else
        MagnetEventHandler = kwargs
    End If
    
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
    Select Case UBound(glf_playerState.Keys())
        Case -1:
            glf_playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            Glf_BcpAddPlayer 1
            glf_currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                glf_playerState.Add "PLAYER 2", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 2
            End If
        Case 1:
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                glf_playerState.Add "PLAYER 3", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 3
            End If     
        Case 2:   
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                glf_playerState.Add "PLAYER 4", Glf_InitNewPlayer()
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
    DispatchPinEvent GLF_GAME_START, Null
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
    DispatchPinEvent "trough_eject", Null
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
    
    DispatchPinEvent GLF_BALL_WILL_END, Null
    DispatchPinEvent GLF_BALL_ENDING, Null
    DispatchPinEvent GLF_BALL_ENDED, Null
    SetPlayerState GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            If UBound(glf_playerState.Keys()) > 0 Then
                glf_currentPlayer = "PLAYER 2"
            End If
        Case "PLAYER 2":
            If UBound(glf_playerState.Keys()) > 1 Then
                glf_currentPlayer = "PLAYER 3"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 3":
            If UBound(glf_playerState.Keys()) > 2 Then
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
        glf_playerState.RemoveAll()
    Else
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000 
    End If
    
End Function

Public Function EndOfBallNextPlayer(args)
    DispatchPinEvent GLF_NEXT_PLAYER, Null
End Function
'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************
Class GlfDebugLogFile
	Private Filename
	Private TxtFileStream

	Public default Function init()
        Filename = cGameName + "_" & GetTimeStamp & "_debug_log.txt"
		TxtFileStream = Null
		Set Init = Me
	End Function

	Public Sub EnableLogs()
		If glf_debugEnabled = True Then
			DisableLogs()
			Dim FormattedMsg, Timestamp, fso, logFolder
			Set fso = CreateObject("Scripting.FileSystemObject")
			logFolder = "glf_logs"
			If Not fso.FolderExists(logFolder) Then
				fso.CreateFolder logFolder
			End If
			Set TxtFileStream = fso.OpenTextFile(logFolder & "\" & Filename, 8, True)
		End If
	End Sub

	Public Sub DisableLogs()
		If Not IsNull(TxtFileStream) Then
			TxtFileStream.Close
		End If
	End Sub
	
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
			Timestamp = GetTimeStamp
			FormattedMsg = Timestamp & ": " & label & ": " & message
			TxtFileStream.WriteLine FormattedMsg
			Debug.Print label & ": " & message
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

Function DispatchQueuePinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        'glf_debugLog.WriteToLog "DispatchRelayPinEvent", e & " has no listeners"
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k,i,retArgs
    Dim handlers : Set handlers = glf_pinEvents(e)
    If IsNull(kwargs) Then
        Set kwargs = GlfKwargs()
    End If
    'glf_debugLog.WriteToLog "DispatchReplayPinEvent", e
    For i=0 to UBound(glf_pinEventsOrder(e))
        k = glf_pinEventsOrder(e)(i)
        'glf_debugLog.WriteToLog "DispatchQueuePinEvent"&e, "key: " & k(1) & ", priority: " & k(0)
        'msgbox "DispatchQueuePinEvent: " & e & " , key: " & k(1) & ", priority: " & k(0)
        'msgbox handlers(k(1))(0)
        'Call the handlers.
        'The handlers might return a waitfor command.
        'If NO wait for command, continue calling handlers.
        'IF wait for command, then AddPinEventListener for the waitfor event. The callback handler needs to be ContinueDispatchQueuePinEvent.
        Set retArgs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
        If retArgs.Exists("wait_for") And i<Ubound(glf_pinEventsOrder(e)) Then
            'pause execution of handlers at index I. 
            AddPinEventListener retArgs("wait_for"), k(1) & "_wait_for", "ContinueDispatchQueuePinEvent", k(0), Array(e, kwargs, i+1)
            Exit For
            'add event listener for the wait_for event.
            'pass in the index and handlers from this.
            'in the handler for resume queue event, process from the index the remaining handlers.
        End If
    Next
    Glf_EventBlocks(e).RemoveAll
End Function


'args Array(3)
' Array(original_event, orignal_kwargs, index)
' wait_for kwargs
' event
Function ContinueDispatchQueuePinEvent(args)
    Dim arrContinue : arrContinue = args(0)
    Dim e : e = arrContinue(0)
    Dim kwargs : kwargs = arrContinue(1)
    Dim idx : idx = arrContinue(2)
    
    If Not glf_pinEvents.Exists(e) Then
        'glf_debugLog.WriteToLog "DispatchRelayPinEvent", e & " has no listeners"
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k,i,retArgs
    Dim handlers : Set handlers = glf_pinEvents(e)
    'glf_debugLog.WriteToLog "DispatchReplayPinEvent", e
    For i=idx to UBound(glf_pinEventsOrder(e))
        k = glf_pinEventsOrder(e)(i)
        'glf_debugLog.WriteToLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)

        'Call the handlers.
        'The handlers might return a waitfor command.
        'If NO wait for command, continue calling handlers.
        'IF wait for command, then AddPinEventListener for the waitfor event. The callback handler needs to be ContinueDispatchQueuePinEvent.
        Set retArgs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
        If retArgs.Exists("wait_for") And i<Ubound(glf_pinEventsOrder(e)) Then
            'pause execution of handlers at index I. 
            AddPinEventListener retArgs("wait_for"), k(1) & "_wait_for", "ContinueDispatchQueuePinEvent", k(0), Array(e, kwargs, i)
            Exit For
            'add event listener for the wait_for event.
            'pass in the index and handlers from this.
            'in the handler for resume queue event, process from the index the remaining handlers.
        End If
    Next
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

    If glf_playerState(glf_currentPlayer).Exists(key)  Then
        GetPlayerState = glf_playerState(glf_currentPlayer)(key)
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

    If glf_playerState.Exists(p) Then
        GetPlayerScore = glf_playerState(p)(SCORE)
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
    If glf_playerState(glf_currentPlayer).Exists(key) Then
        prevValue = glf_playerState(glf_currentPlayer)(key)
        glf_playerState(glf_currentPlayer).Remove key
    End If
    glf_playerState(glf_currentPlayer).Add key, value
    
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
    For Each key in glf_playerState(glf_currentPlayer).Keys()
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
Sub Drain_Hit
	UpdateTrough
    If glf_gameStarted = True Then
        glf_BIP = glf_BIP - 1
        DispatchRelayPinEvent GLF_BALL_DRAIN, 1
    End If
End Sub
Sub Drain_UnHit
	UpdateTrough
End Sub

Sub UpdateTrough
	SetDelay "update_trough", "UpdateTroughDebounced", Null, 100
End Sub

Sub UpdateTroughDebounced(args)
	If glf_troughSize > 1 Then
		If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	End If
	If glf_troughSize > 2 Then
		If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	End If
	If glf_troughSize > 3 Then 
		If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	End If
	If glf_troughSize > 4 Then 
		If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	End If
	If glf_troughSize > 5 Then
		If swTrough5.BallCntOver = 0 Then swTrough6.kick 57, 10
	End If
	If glf_troughSize > 6 Then
		If swTrough6.BallCntOver = 0 Then swTrough7.kick 57, 10
	End If

	If glf_lastTroughSw.BallCntOver = 0 Then Drain.kick 57, 10
End Sub
