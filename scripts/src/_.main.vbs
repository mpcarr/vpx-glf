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
Dim glf_diverters : Set glf_diverters = CreateObject("Scripting.Dictionary")
Dim glf_flippers : Set glf_flippers = CreateObject("Scripting.Dictionary")
Dim glf_autofiredevices : Set glf_autofiredevices = CreateObject("Scripting.Dictionary")
Dim glf_ball_holds : Set glf_ball_holds = CreateObject("Scripting.Dictionary")
Dim glf_magnets : Set glf_magnets = CreateObject("Scripting.Dictionary")
Dim glf_segment_displays : Set glf_segment_displays = CreateObject("Scripting.Dictionary")
Dim glf_droptargets : Set glf_droptargets = CreateObject("Scripting.Dictionary")
Dim glf_multiball_locks : Set glf_multiball_locks = CreateObject("Scripting.Dictionary")
Dim glf_multiballs : Set glf_multiballs = CreateObject("Scripting.Dictionary")
Dim glf_shows : Set glf_shows = CreateObject("Scripting.Dictionary")



Dim bcpController : bcpController = Null
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = ""
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0
Dim glf_SeqCount : glf_SeqCount = 0

Dim glf_ballsPerGame : glf_ballsPerGame = 3
Dim glf_troughSize : glf_troughSize = tnob
Dim glf_lastTroughSw : glf_lastTroughSw = Null

Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False
Dim glf_debug_level : glf_debug_level = "Info"

Glf_RegisterLights()
Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController
    Set bcpController = (new GlfVpxBcpController)(bcpPort, bcpExeName)
End Sub

Public Sub Glf_WriteDebugLog(name, message)
	If glf_debug_level = "Debug" Then
		glf_debugLog.WriteToLog name, message
	End If
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
		Dim switchNumber : switchNumber = 1
		Dim lightsNumber : lightsNumber = 1
		Dim switchesYaml : switchesYaml = "#config_version=6" & vbCrLf & vbCrLf
		Dim lightsYaml : lightsYaml = "#config_version=6" & vbCrLf & vbCrLf
		lightsYaml = lightsYaml + "lights:" & vbCrLf
		Dim monitorYaml : monitorYaml = "light:" & vbCrLf
		Dim godotLightScene : godotLightScene = ""
		For Each light in glf_lights
			monitorYaml = monitorYaml + "  " & light.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    size: 0.04" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& light.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& light.y/tableheight & vbCrLf

			lightsYaml = lightsYaml + "  " & light.name & ":"&vbCrLf
			lightsYaml = lightsYaml + "    number: " & lightsNumber & vbCrLf
			lightsYaml = lightsYaml + "    subtype: led" & vbCrLf
			lightsYaml = lightsYaml + "    type: rgb" & vbCrLf
			lightsYaml = lightsYaml + "    tags: " & light.BlinkPattern & vbCrLf
			lightsNumber = lightsNumber + 1

			godotLightScene = godotLightScene + "[node name="""&light.name&""" type=""Sprite2D"" parent=""lights""]" & vbCrLf
			godotLightScene = godotLightScene + "position = Vector2("&light.x*scaleFactor&", "&light.y*scaleFactor&")" & vbCrLf
			godotLightScene = godotLightScene + "script = ExtResource(""3_qb2nn"")" & vbCrLf

			Dim splitTagArray : splitTagArray = Split(light.BlinkPattern, ",")
			Dim outputTagString : outputTagString = ""
			Dim i
			For i = LBound(splitTagArray) To UBound(splitTagArray)
				outputTagString = outputTagString & """" & Trim(splitTagArray(i)) & """"
				If i < UBound(splitTagArray) Then
					outputTagString = outputTagString & ", "
				End If
			Next

			godotLightScene = godotLightScene + "tags = ["&outputTagString&"]" & vbCrLf
			godotLightScene = godotLightScene + vbCrLf
		Next

		monitorYaml = monitorYaml + vbCrLf
		monitorYaml = monitorYaml + "switch:" & vbCrLf
		switchesYaml = switchesYaml + "switches:" & vbCrLf
		
		For Each switch in glf_switches
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& switch.y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  " & switch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		For Each switch in glf_spinners
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& switch.y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  " & switch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		For Each switch in glf_slingshots
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.BlendDisableLighting & vbCrLf
			monitorYaml = monitorYaml + "    y: "& 1-switch.BlendDisableLightingFromBelow & vbCrLf
			switchesYaml = switchesYaml + "  " & switch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		Dim troughCount
		For troughCount=1 to tnob
			monitorYaml = monitorYaml + "  s_trough" & troughCount & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& Eval("swTrough"&troughCount).x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& Eval("swTrough"&troughCount).y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  s_trough" & troughCount & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		switchesYaml = switchesYaml + "  s_trough_jam" & ":"&vbCrLf
		switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
		switchesYaml = switchesYaml + "    tags: " & vbCrLf
		switchNumber = switchNumber + 1

		monitorYaml = monitorYaml + "  s_start:"&vbCrLf
		monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
		monitorYaml = monitorYaml + "    x: 0.95" & vbCrLf
		monitorYaml = monitorYaml + "    y: 0.95" & vbCrLf
		switchesYaml = switchesYaml + "  s_start:"&vbCrLf
		switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
		switchesYaml = switchesYaml + "    tags: start" & vbCrLf
		switchNumber = switchNumber + 1


		Dim fso, modesFolder, TxtFileStream, monitorFolder, configFolder
		Set fso = CreateObject("Scripting.FileSystemObject")
		monitorFolder = "glf_mpf\monitor\"
		configFolder = "glf_mpf\config\"
		If Not fso.FolderExists("glf_mpf") Then
			fso.CreateFolder "glf_mpf"
		End If
		If Not fso.FolderExists("glf_mpf\monitor") Then
			fso.CreateFolder "glf_mpf\monitor"
		End If
		If Not fso.FolderExists("glf_mpf\config") Then
			fso.CreateFolder "glf_mpf\config"
		End If
		Set TxtFileStream = fso.OpenTextFile(monitorFolder & "\monitor.yaml", 2, True)
		TxtFileStream.WriteLine monitorYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\switches.yaml", 2, True)
		TxtFileStream.WriteLine switchesYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\lights.yaml", 2, True)
		TxtFileStream.WriteLine lightsYaml
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
		Glf_WriteDebugLog "Init", mode.Name
		If Not IsNull(mode.ShowPlayer) Then
			With mode.ShowPlayer()
				Dim show_settings
				For Each show_settings in .EventShows()
					If Not IsNull(show_settings.Show) And show_settings.Action = "play" Then
						show_settings.InternalCacheId = CStr(show_count)
						show_count = show_count + 1
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
		Dim x,state
		If UBound(mode.ModeShots) > -1 Then
			Dim mode_shot
			For Each mode_shot in mode.ModeShots
				Dim shot_profile : Set shot_profile = Glf_ShotProfiles(mode_shot.Profile)
				
				If mode_shot.InternalCacheId = -1 Then
					mode_shot.InternalCacheId = shot_count
					shot_count = shot_count + 1
				End If
				For x=0 to shot_profile.StatesCount
					Set state = shot_profile.StateForIndex(x)
					If state.InternalCacheId = -1 Then
						state.InternalCacheId = CStr(show_count)
						show_count = show_count + 1
					End If

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

		If UBound(mode.ModeStateMachines) > -1 Then
			Dim mode_state_machine,state_count
			state_count = 0
			For Each mode_state_machine in mode.ModeStateMachines
				
				For x=0 to UBound(mode_state_machine.StateItems)
					Set state = mode_state_machine.StateItems()(x)
					If state.InternalCacheId = -1 Then
						state.InternalCacheId = CStr(state_count)
						state_count = state_count + 1
					End If
					If Not IsNull(state.ShowWhenActive().Show) Then
						If state.ShowWhenActive().Action = "play" Then
							state.ShowWhenActive().InternalCacheId = CStr(show_count)
							show_count = show_count + 1
							cached_show = Glf_ConvertShow(state.ShowWhenActive().Show, state.ShowWhenActive().Tokens)
							glf_cached_shows.Add mode.name & "_" & mode_state_machine.Name & "_" & state.Name & "_" & state.ShowWhenActive().Key & "_" & state.InternalCacheId & "_" & state.ShowWhenActive().InternalCacheId, cached_show
						End If
					End If
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

	Dim glfDebugLevel : glfDebugLevel = Table1.Option("Glf Debug Log Level", 0, 1, 1, 0, 0, Array("Info", "Debug"))
	If glfDebugLevel = 1 Then
		glf_debug_level = "Debug"
	Else
		glf_debug_level = "Info"
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
			
			tag = "T_" & Trim(tag)
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
				tag = "T_" & Trim(tag)
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
			tmp = Glf_ReplaceAnyPlayerAttributes(tmp)
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
	Dim parts : parts = Split(value, ":")
	Dim event_delay : event_delay = 0
	If UBound(parts) = 1 Then
		value = parts(0)
		event_delay= parts(1)
	End If


	Dim condition : condition = Glf_IsCondition(value)
	If IsNull(condition) Then
		Glf_ParseEventInput = Array(value, value, Null, event_delay)
	Else
		dim conditionReplaced : conditionReplaced = Glf_ReplaceCurrentPlayerAttributes(condition)
		conditionReplaced = Glf_ReplaceAnyPlayerAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceDeviceAttributes(conditionReplaced)
		templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
		templateCode = templateCode & vbTab & "On Error Resume Next" & vbCrLf
		templateCode = templateCode & vbTab & Glf_ConvertCondition(conditionReplaced, "Glf_" & glf_FuncCount) & vbCrLf
		templateCode = templateCode & vbTab & "If Err Then Glf_" & glf_FuncCount & " = False" & vbCrLf
		templateCode = templateCode & "End Function"
		'msgbox templateCode
		ExecuteGlobal templateCode
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		glf_FuncCount = glf_FuncCount + 1
		Glf_ParseEventInput = Array(Replace(value, "{"&condition&"}", funcRef) ,Replace(value, "{"&condition&"}", ""), funcRef, event_delay)
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

Function Glf_ReplaceAnyPlayerAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "players\[([0-3]+)\]\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    replacement = "GetPlayerStateForPlayer($1, ""$2"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceAnyPlayerAttributes = outputString
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

Function Glf_CheckForGetPlayerState(inputString)
    Dim pattern, regex, matches, match, hasGetPlayerState, attribute, playerNumber
	inputString = Glf_ReplaceCurrentPlayerAttributes(inputString)
	inputString = Glf_ReplaceAnyPlayerAttributes(inputString)
    pattern = "GetPlayerState\(""([a-zA-Z0-9_]+)""\)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False
    
    Set matches = regex.Execute(inputString)
    If matches.Count > 0 Then
        hasGetPlayerState = True
		playerNumber = -1 'Current Player
        attribute = matches(0).SubMatches(0)
    Else
        hasGetPlayerState = False
        attribute = ""
		playerNumber = Null
    End If


	pattern = "GetPlayerStateForPlayer\(([0-3]), ""([a-zA-Z0-9_]+)""\)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False
    
    Set matches = regex.Execute(inputString)
    If matches.Count > 0 Then
        hasGetPlayerState = True
		playerNumber = Int(matches(0).SubMatches(0))
        attribute = matches(0).SubMatches(1)
    Else
        hasGetPlayerState = False
        attribute = ""
		playerNumber = Null
    End If

    Set regex = Nothing
    Set matches = Nothing
    
    Glf_CheckForGetPlayerState = Array(hasGetPlayerState, attribute, playerNumber)
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
	isVariable = Glf_IsCondition(falsePart)
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
    Dim padChar, width, result, align, hasCommas
	
	If CStr(value) = "False" Then
		Glf_FormatValue = ""
		Exit Function
	End If

    ' Default values
    padChar = " " ' Default padding character is space
    align = ">"   ' Default alignment is right
    width = 0     ' Default width is 0 (no padding)
    hasCommas = False ' Default: No thousand separators

    ' Check for :, in the format string
    If InStr(formatString, ",") > 0 Then
        hasCommas = True
        formatString = Replace(formatString, ",", "") ' Remove , from format string
    End If

    ' Parse the remaining format string
    If Len(formatString) >= 2 Then
        padChar = Mid(formatString, 1, 1)
        align = Mid(formatString, 2, 1)
        width = CInt(Mid(formatString, 3))
    End If

    ' Format the value
    If hasCommas And IsNumeric(value) Then
        ' Add commas as thousand separators
        Dim numStr, decimalPart
        numStr = CStr(value)
        If InStr(numStr, ".") > 0 Then
            decimalPart = Mid(numStr, InStr(numStr, "."))
            numStr = Left(numStr, InStr(numStr, ".") - 1)
        Else
            decimalPart = ""
        End If

        Dim i, formattedNum
        formattedNum = ""
        For i = Len(numStr) To 1 Step -1
            formattedNum = Mid(numStr, i, 1) & formattedNum
            If ((Len(numStr) - i) Mod 3 = 2) And (i > 1) Then
                formattedNum = "," & formattedNum
            End If
        Next
        value = formattedNum & decimalPart
    End If

    ' Apply alignment and padding
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


Function glf_ReadShowYAMLFiles()
    Dim fso, folder, file, yamlFiles, fileContent
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the directory exists
    If Not fso.FolderExists(directoryPath) Then
        WScript.Echo "Directory does not exist: " & directoryPath
        Exit Function
    End If
    
    ' Initialize the array to store file contents
    ReDim yamlFiles(-1)
    
    ' Get the folder object
    Set folder = fso.GetFolder(directoryPath)
    
    ' Iterate through the files in the directory
    For Each file In folder.Files
        ' Check if the file has a .yaml extension
        If LCase(fso.GetExtensionName(file.Name)) = "yaml" Then
            ' Read the file content
            Set fileContent = fso.OpenTextFile(file.Path, 1) ' 1 = ForReading
            ReDim Preserve yamlFiles(UBound(yamlFiles) + 1)
            yamlFiles(UBound(yamlFiles)) = fileContent.ReadAll
            fileContent.Close
        End If
    Next
    
    ' Return the array of YAML file contents
    ReadYAMLFiles = yamlFiles
End Function

Sub glf_ConvertYamlShowToGlfShow(yamlFilePath)
    Dim fso, file, content, lines, line, output, i, stepLights
    Dim glf_ShowName, stepTime, lightsDict, key, lightName, color, intensity, lightParts
    
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
        ElseIf InStr(line, ":") > 0 Then
            lightParts = Split(line, ":")
			key = lightParts(0)
            lightName = Trim(key)
			
            color = lightParts(1)
			color = Trim(color)
			color = Replace(color, """", "")
            'msgbox key & "<>" & color
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
    msgbox Len(stepLights)
    ' Close the final step and the show
	If Len(stepLights) = 0 Then
		output = output & vbTab & vbTab & ".Lights = Array()" & vbCrLf
	Else
		output = output & vbTab & vbTab & ".Lights = Array(" & _ 
		SplitStringWithUnderscore(Left(stepLights, Len(stepLights) - 1), 1500) & ")" & vbCrLf
	End If
	
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
					tagLights = glf_lightTags("T_"&lightParts(0)).Keys()
					lightsCount = UBound(tagLights)+1
				Else
					If IsNull(token) Then
						lightsCount = lightsCount + 1
					Else
						'resolve token lights
						If Not glf_lightNames.Exists(tokens(token)) Then
							'token is a tag
							tagLights = glf_lightTags("T_"&tokens(token)).Keys()
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
					tagLights = glf_lightTags("T_"&lightParts(0)).Keys()
					For Each tagLight in tagLights
						If UBound(lightParts) >=1 Then
							seqArray(x) = tagLight & "|"&lightParts(1)&"|" & AdjustHexColor(lightColor, lightParts(1))
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
							seqArray(x) = lightParts(0) & "|"&lightParts(1)&"|"&AdjustHexColor(lightColor, lightParts(1))
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
							tagLights = glf_lightTags("T_"&tokens(token)).Keys()
							For Each tagLight in tagLights
								If UBound(lightParts) >=1 Then
									seqArray(x) = tagLight & "|"&lightParts(1)&"|"&AdjustHexColor(lightColor, lightParts(1))
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
								seqArray(x) = tokens(token) & "|"&lightParts(1)&"|"&AdjustHexColor(lightColor, lightParts(1))
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
		'Glf_WriteDebugLog "Convert Show", Join(seqArray)
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
	Private m_raw, m_value, m_isGetRef, m_isPlayerState, m_playerStateValue, m_playerStatePlayer
  
    Public Property Get Value() 
		If m_isGetRef = True Then
			Value = GetRef(m_value)()
		Else
			Value = m_value
		End If
	End Property

    Public Property Get Raw() : Raw = m_raw : End Property

	Public Property Get IsPlayerState() : IsPlayerState = m_isPlayerState : End Property
	Public Property Get PlayerStateValue() : PlayerStateValue = m_playerStateValue : End Property		
	Public Property Get PlayerStatePlayer() : PlayerStatePlayer = m_playerStatePlayer : End Property		


	Public default Function init(input)
        m_raw = input
        Dim parsedInput : parsedInput = Glf_ParseInput(input)
		Dim playerState : playerState = Glf_CheckForGetPlayerState(input)
		m_isPlayerState = playerState(0)
		m_playerStateValue = playerState(1)
		m_playerStatePlayer = playerState(2)
        m_value = parsedInput(0)
        m_isGetRef = parsedInput(2)
	    Set Init = Me
	End Function

End Class

Function Glf_FadeRGB(light, color1, color2, steps)

	Dim r1, g1, b1, r2, g2, b2
	Dim i
	Dim r, g, b
	color1 = clng( RGB( Glf_HexToInt(Left(color1, 2)), Glf_HexToInt(Mid(color1, 3, 2)), Glf_HexToInt(Right(color1, 2)))  )
	color2 = clng( RGB( Glf_HexToInt(Left(color2, 2)), Glf_HexToInt(Mid(color2, 3, 2)), Glf_HexToInt(Right(color2, 2)))  )
	
	r1 = color1 Mod 256
	g1 = (color1 \ 256) Mod 256
	b1 = (color1 \ (256 * 256)) Mod 256

	r2 = color2 Mod 256
	g2 = (color2 \ 256) Mod 256
	b2 = (color2 \ (256 * 256)) Mod 256

	ReDim outputArray(steps - 1)
	For i = 0 To steps - 1
		r = r1 + (r2 - r1) * i / (steps - 1)
		g = g1 + (g2 - g1) * i / (steps - 1)
		b = b1 + (b2 - b1) * i / (steps - 1)
		outputArray(i) = light & "|100|" & Glf_RGBToHex(CInt(r), CInt(g), CInt(b))
	Next
	Glf_FadeRGB = outputArray
End Function

Function Glf_RGBToHex(r, g, b)
	Glf_RGBToHex = Right("0" & Hex(r), 2) & _
	Right("0" & Hex(g), 2) & _
	Right("0" & Hex(b), 2)
End Function

Private Function Glf_HexToInt(hex)
	Glf_HexToInt = CInt("&H" & hex)
End Function

'******************************************************
'*****   GLF Shows 		                           ****
'******************************************************

Function CreateGlfShow(name)
	Dim show : Set show = (new GlfShow)(name)
	'msgbox name
	glf_shows.Add name, show
	Set CreateGlfShow = show
End Function

With CreateGlfShow("on")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100")
	End With
End With

With CreateGlfShow("off")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

With CreateGlfShow("flash")
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

With CreateGlfShow("flash_color")
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|(color)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

With CreateGlfShow("led_color")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|(color)")
	End With
End With

With GlfShotProfiles("default")
	With .States("on")
		.Show = "flash"
	End With
	With .States("off")
		.Show = "off"
	End With
End With

With GlfShotProfiles("flash_color")
	With .States("off")
		.Show = "off"
	End With
	With .States("on")
		.Show = "flash_color"
	End With	
End With

Function AdjustHexColor(hexColor, percentage)
    ' Ensure percentage is between 0 and 100
    If percentage < 0 Then percentage = 0
    If percentage > 100 Then percentage = 100

    ' Parse the R, G, B components
    Dim r, g, b
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))

    ' Adjust the RGB components by the percentage
    r = Int(r * (percentage / 100))
    g = Int(g * (percentage / 100))
    b = Int(b * (percentage / 100))

    ' Ensure the values are within the valid range (0 to 255)
    r = FixRange(r)
    g = FixRange(g)
    b = FixRange(b)

    ' Convert back to hex and return the adjusted color
    AdjustHexColor = PadHex(Hex(r)) & PadHex(Hex(g)) & PadHex(Hex(b))
End Function

' Helper function to ensure values stay within range
Function FixRange(value)
    If value < 0 Then value = 0
    If value > 255 Then value = 255
    FixRange = value
End Function

' Helper function to pad single digit hex values with a leading zero
Function PadHex(hexValue)
    If Len(hexValue) < 2 Then
        PadHex = "0" & hexValue
    Else
        PadHex = hexValue
    End If
End Function


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
Const GLF_CURRENT_BALL = "ball"
Const GLF_INITIALS = "initials"

