'VPX Game Logic Framework (https://mpcarr.github.io/vpx-glf/)

'
Dim glf_currentPlayer : glf_currentPlayer = Null
Dim glf_canAddPlayers : glf_canAddPlayers = True
Dim glf_PI : glf_PI = 4 * Atn(1)
Dim glf_plunger, glf_ballsearch, glf_highscore
glf_ballsearch = Null
glf_highscore = Null
Dim glf_ballsearch_enabled : glf_ballsearch_enabled = False
Dim glf_gameStarted : glf_gameStarted = False
Dim glf_gameTilted : glf_gameTilted = False
Dim glf_gameEnding : glf_gameEnding = False
Dim glf_last_switch_hit_time : glf_last_switch_hit_time = 0
Dim glf_last_ballsearch_reset_time : glf_last_ballsearch_reset_time = 0
Dim glf_last_switch_hit : glf_last_switch_hit = ""
Dim glf_current_virtual_tilt : glf_current_virtual_tilt = 0
Dim glf_tilt_sensitivity : glf_tilt_sensitivity = 7
Dim glf_flex_alphadmd : Set glf_flex_alphadmd = Nothing
Dim glf_flex_alphadmd_enabled : glf_flex_alphadmd_enabled = False
Dim glf_flex_alphadmd_segments(31)
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
Dim glf_cached_rgb_fades : Set glf_cached_rgb_fades = CreateObject("Scripting.Dictionary")
Dim glf_lightPriority : Set glf_lightPriority = CreateObject("Scripting.Dictionary")
Dim glf_lightColorLookup : Set glf_lightColorLookup = CreateObject("Scripting.Dictionary")
Dim glf_lightMaps : Set glf_lightMaps = CreateObject("Scripting.Dictionary")
Dim glf_lightStacks : Set glf_lightStacks = CreateObject("Scripting.Dictionary")
Dim glf_lightTags : Set glf_lightTags = CreateObject("Scripting.Dictionary")
Dim glf_lightNames : Set glf_lightNames = CreateObject("Scripting.Dictionary")
Dim glf_modes : Set glf_modes = CreateObject("Scripting.Dictionary")
Dim glf_timers : Set glf_timers = CreateObject("Scripting.Dictionary")
Dim glf_codestr : glf_codestr = ""
Dim glf_state_machines : Set glf_state_machines = CreateObject("Scripting.Dictionary")
Dim glf_ball_devices : Set glf_ball_devices = CreateObject("Scripting.Dictionary")
Dim glf_diverters : Set glf_diverters = CreateObject("Scripting.Dictionary")
Dim glf_flippers : Set glf_flippers = CreateObject("Scripting.Dictionary")
Dim glf_autofiredevices : Set glf_autofiredevices = CreateObject("Scripting.Dictionary")
Dim glf_ball_holds : Set glf_ball_holds = CreateObject("Scripting.Dictionary")
Dim glf_magnets : Set glf_magnets = CreateObject("Scripting.Dictionary")
Dim glf_segment_displays : Set glf_segment_displays = CreateObject("Scripting.Dictionary")
Dim glf_drop_targets : Set glf_drop_targets = CreateObject("Scripting.Dictionary")
Dim glf_standup_targets : Set glf_standup_targets = CreateObject("Scripting.Dictionary")
Dim glf_multiball_locks : Set glf_multiball_locks = CreateObject("Scripting.Dictionary")
Dim glf_multiballs : Set glf_multiballs = CreateObject("Scripting.Dictionary")
Dim glf_shows : Set glf_shows = CreateObject("Scripting.Dictionary")
Dim glf_initialVars : Set glf_initialVars = CreateObject("Scripting.Dictionary")
Dim glf_dispatch_await : Set glf_dispatch_await = CreateObject("Scripting.Dictionary")
Dim glf_dispatch_handlers_await : Set glf_dispatch_handlers_await = CreateObject("Scripting.Dictionary")
Dim glf_dispatch_current_kwargs
Dim glf_dispatch_lightmaps_await : Set glf_dispatch_lightmaps_await = CreateObject("Scripting.Dictionary")
Dim glf_machine_vars : Set glf_machine_vars = CreateObject("Scripting.Dictionary")
Dim glf_achievements : Set glf_achievements = CreateObject("Scripting.Dictionary")
Dim glf_sound_buses : Set glf_sound_buses = CreateObject("Scripting.Dictionary")
Dim glf_sounds : Set glf_sounds = CreateObject("Scripting.Dictionary")
Dim glf_combo_switches : Set glf_combo_switches = CreateObject("Scripting.Dictionary")
Dim glf_funcRefMap : Set glf_funcRefMap = CreateObject("Scripting.Dictionary")
Dim bcpController : bcpController = Null
Dim glf_debugBcpController : glf_debugBcpController = Null
Dim glf_hasDebugController : glf_hasDebugController = False
Dim glf_monitor_player_state : glf_monitor_player_state = ""
Dim glf_monitor_modes : glf_monitor_modes = ""
Dim glf_monitor_event_stream : glf_monitor_event_stream = ""
Dim glf_running_modes : glf_running_modes = ""
Dim glf_production_mode : glf_production_mode = False
Dim useGlfBCPMonitor : useGlfBCPMonitor = False
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = CGameName & "_gmc.exe"
Dim glf_monitor_player_vars : glf_monitor_player_vars = false
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0
Dim glf_SeqCount : glf_SeqCount = 0
Dim glf_max_dispatch : glf_max_dispatch = 25
Dim glf_max_lightmap_sync : glf_max_lightmap_sync = -1
Dim glf_max_lightmap_sync_enabled : glf_max_lightmap_sync_enabled = False
Dim glf_max_lights_test : glf_max_lights_test = 0

Dim glf_master_volume : glf_master_volume = 0.8

Dim glf_table

Dim glf_troughSize : glf_troughSize = tnob
Dim glf_lastTroughSw : glf_lastTroughSw = Null
Dim glf_game

Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False
Dim glf_debug_level : glf_debug_level = "Info"


Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController(args)
	If glf_production_mode = True Then
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(bcpExeName) Then
			Set bcpController = (new GlfVpxBcpController)(bcpPort, bcpExeName)	
		Else
			MsgBox "Missing GMCDisplay file"
		End If
	Else
		Set bcpController = (new GlfVpxBcpController)(bcpPort, "")
	End If
End Sub

Public Sub Glf_ConnectToDebugBCPMediaController(args)
    Set glf_debugBcpController = (new GlfMonitorBcpController)(5051, "glf_monitor.exe")
	glf_hasDebugController = True
End Sub

Public Sub Glf_WriteDebugLog(name, message)
	If glf_debug_level = "Debug" Then
		glf_debugLog.WriteToLog name, message
		'Glf_MonitorEventStream name, message
	End If
End Sub

Public Function SwitchHandler(handler, args)
	SwitchHandler = False
	Select Case handler
		Case "BaseModeDeviceEventHandler"
			BaseModeDeviceEventHandler args
			SwitchHandler = True
	End Select

End Function

Public Sub Glf_Init(ByRef table)
    Set glf_table = table
	With GlfGameSettings()
		.BallsPerGame = 3
	End With
	Glf_Options Null 'Force Options Check
	Glf_RegisterLights()
	glf_debugLog.WriteToLog "Init", "Start"
	If glf_troughSize > 0 Then : swTrough1.DestroyBall : Set glf_ball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1) : Set glf_lastTroughSw = swTrough1 : End If
	If glf_troughSize > 1 Then : swTrough2.DestroyBall : Set glf_ball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2) : Set glf_lastTroughSw = swTrough2 : End If
	If glf_troughSize > 2 Then : swTrough3.DestroyBall : Set glf_ball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3) : Set glf_lastTroughSw = swTrough3 : End If
	If glf_troughSize > 3 Then : swTrough4.DestroyBall : Set glf_ball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4) : Set glf_lastTroughSw = swTrough4 : End If
	If glf_troughSize > 4 Then : swTrough5.DestroyBall : Set glf_ball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5) : Set glf_lastTroughSw = swTrough5 : End If
	If glf_troughSize > 5 Then : swTrough6.DestroyBall : Set glf_ball6 = swTrough6.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6) : Set glf_lastTroughSw = swTrough6 : End If
	If glf_troughSize > 6 Then : swTrough7.DestroyBall : Set glf_ball7 = swTrough7.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7) : Set glf_lastTroughSw = swTrough7 : End If
	If glf_troughSize > 7 Then : Drain.DestroyBall : Set glf_ball8 = Drain.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8) : End If
	
	
    Dim codestr : codestr = ""
	Dim switch
	For Each switch in Glf_Switches
		codestr = codestr & "Sub " & switch.Name & "_Hit() : If Not glf_gameTilted Then : DispatchPinEvent """ & switch.Name & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& switch.Name &""": End If : End Sub" & vbCrLf
		codestr = codestr & "Sub " & switch.Name & "_UnHit() : If Not glf_gameTilted Then : DispatchPinEvent """ & switch.Name & "_inactive"", ActiveBall : End If  : End Sub" & vbCrLf
	Next
	
    codestr = codestr & vbCrLf

	Dim slingshot
	For Each slingshot in Glf_Slingshots
		codestr = codestr & "Sub " & slingshot.Name & "_Slingshot() : If Not glf_gameTilted Then : DispatchPinEvent """ & slingshot.Name & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& slingshot.Name &""": End If  : End Sub" & vbCrLf
	Next
	
    codestr = codestr & vbCrLf

	Dim spinner
	For Each spinner in Glf_Spinners
		codestr = codestr & "Sub " & spinner.Name & "_Spin() : If Not glf_gameTilted Then : DispatchPinEvent """ & spinner.Name & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& spinner.Name &""": End If  : End Sub" & vbCrLf
	Next

	Dim drop_target
	Dim drop_array, using_roth_drops
	using_roth_drops = False
	drop_array = Array()
	For Each drop_target in glf_drop_targets.Items()
		codestr = codestr & "Sub " & drop_target.Switch & "_Hit() : If Not glf_gameTilted Then : If glf_drop_targets(""" & drop_target.Name & """).UseRothDroptarget = True Then : DTHit glf_drop_targets(""" & drop_target.Name & """).RothDTSwitchID : Else : DispatchPinEvent """ & drop_target.Switch & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& drop_target.Switch &""": End If : End If : End Sub" & vbCrLf
		codestr = codestr & "Sub " & drop_target.Switch & "_UnHit() : If Not glf_gameTilted Then : If glf_drop_targets(""" & drop_target.Name & """).UseRothDroptarget = False Then : DispatchPinEvent """ & drop_target.Switch & "_inactive"", ActiveBall : End If : End If : End Sub" & vbCrLf
	Next

	Dim standup_target
	For Each standup_target in glf_standup_targets.Items()
		codestr = codestr & "Sub " & standup_target.Switch & "_Hit() : If Not glf_gameTilted Then : If glf_standup_targets(""" & standup_target.Name & """).UseRothStanduptarget = True Then : STHit glf_standup_targets(""" & standup_target.Name & """).RothSTSwitchID : Else : DispatchPinEvent """ & standup_target.Switch & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& standup_target.Switch &""": End If : End If : End Sub" & vbCrLf
		codestr = codestr & "Sub " & standup_target.Switch & "_UnHit() : If Not glf_gameTilted Then : If glf_standup_targets(""" & standup_target.Name & """).UseRothStanduptarget = False Then : DispatchPinEvent """ & standup_target.Switch & "_inactive"", ActiveBall : End If : End If : End Sub" & vbCrLf
	Next
	
    codestr = codestr & vbCrLf

	ExecuteGlobal codestr
	
	If glf_debugEnabled = True Then

		'***GLFMPF_EXPORT_START***
		glf_debugLog.WriteToLog "Init", "Exporting MPF Config"
		' Calculate the scale factor
		Dim scaleFactor
		scaleFactor = 1080 / tableheight

		Dim light
		Dim switchNumber : switchNumber = 0
		Dim lightsNumber : lightsNumber = 0
		Dim coilsNumber : coilsNumber = 0
		Dim switchesYaml : switchesYaml = "#config_version=6" & vbCrLf & vbCrLf
		Dim coilsYaml : coilsYaml = "#config_version=6" & vbCrLf & vbCrLf
		coilsYaml = coilsYaml + "coils:" & vbCrLf
		Dim shotProfilesYaml : shotProfilesYaml = "#config_version=6" & vbCrLf & vbCrLf
		shotProfilesYaml = shotProfilesYaml + "shot_profiles:" & vbCrLf
		Dim playerVarsYaml : playerVarsYaml = "#config_version=6" & vbCrLf & vbCrLf
		playerVarsYaml = playerVarsYaml + "player_vars:" & vbCrLf
		Dim ballDevicesYaml : ballDevicesYaml = "#config_version=6" & vbCrLf & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "ball_devices:" & vbCrLf
		Dim configYaml : configYaml = "#config_version=6" & vbCrLf & vbCrLf
		configYaml = configYaml + "config:" & vbCrLf
		configYaml = configYaml + "  - lights.yaml" & vbCrLf
		configYaml = configYaml + "  - ball_devices.yaml" & vbCrLf
		configYaml = configYaml + "  - coils.yaml" & vbCrLf
		configYaml = configYaml + "  - switches.yaml" & vbCrLf
		configYaml = configYaml + vbCrLf
		configYaml = configYaml + "playfields:" & vbCrLf
		configYaml = configYaml + "  playfield:" & vbCrLf
		configYaml = configYaml + "    tags: default" & vbCrLf
		configYaml = configYaml + "    default_source_device: balldevice_plunger" & vbCrLf

		Dim lightsYaml : lightsYaml = "#config_version=6" & vbCrLf & vbCrLf
		lightsYaml = lightsYaml + "lights:" & vbCrLf
		Dim monitorYaml : monitorYaml = "lights:" & vbCrLf
		Dim godotLightScene : godotLightScene = ""
		For Each light in glf_lights
			monitorYaml = monitorYaml + "  " & light.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    size: 0.04" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& light.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& light.y/tableheight & vbCrLf

			lightsYaml = lightsYaml + "  " & light.name & ":"&vbCrLf
			lightsYaml = lightsYaml + "    number: " & lightsNumber & vbCrLf
			lightsYaml = lightsYaml + "    subtype: led" & vbCrLf
			lightsYaml = lightsYaml + "    size: 0.04" & vbCrLf
			lightsYaml = lightsYaml + "    type: rgb" & vbCrLf
			lightsYaml = lightsYaml + "    tags: " & light.BlinkPattern & vbCrLf
			lightsYaml = lightsYaml + "    x: "& light.x/tablewidth & vbCrLf
			lightsYaml = lightsYaml + "    y: "& light.y/tableheight & vbCrLf
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
		monitorYaml = monitorYaml + "switches:" & vbCrLf
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
			switchesYaml = switchesYaml + "    x: "& switch.x/tablewidth & vbCrLf
			switchesYaml = switchesYaml + "    y: "& switch.y/tableheight & vbCrLf
			switchNumber = switchNumber + 1
		Next
		Dim vpxSwitch
		For Each switch in glf_standup_targets.Items()
			Set vpxSwitch = Eval(switch.switch)
			monitorYaml = monitorYaml + "  " & vpxSwitch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& vpxSwitch.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& vpxSwitch.y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  " & vpxSwitch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchesYaml = switchesYaml + "    x: "& vpxSwitch.x/tablewidth & vbCrLf
			switchesYaml = switchesYaml + "    y: "& vpxSwitch.y/tableheight & vbCrLf
			switchNumber = switchNumber + 1
		Next
		For Each switch in glf_drop_targets.Items()
			Set vpxSwitch = Eval(switch.switch)
			monitorYaml = monitorYaml + "  " & vpxSwitch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: -1" & vbCrLf
			monitorYaml = monitorYaml + "    y: -1" & vbCrLf
			switchesYaml = switchesYaml + "  " & vpxSwitch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchesYaml = switchesYaml + "    x: -1" & vbCrLf
			switchesYaml = switchesYaml + "    y: -1" & vbCrLf
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
		Dim troughCount, troughSwitchesArr()
		ReDim troughSwitchesArr(tnob)
		configYaml = configYaml + vbCrLf & "virtual_platform_start_active_switches:" & vbCrLf
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
			troughSwitchesArr(troughCount-1) = "s_trough" & troughCount
			configYaml = configYaml + "  - s_trough" & troughCount & vbCrLf
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

		dim device

		ballDevicesYaml = ballDevicesYaml + "  bd_trough:" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    ball_switches: "&Join(troughSwitchesArr, ",")&" s_trough_jam" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    eject_coil: c_trough_eject" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    tags: trough, home, drain" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    jam_switch: s_trough_jam" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    eject_targets: balldevice_plunger" & vbCrLf
		

		coilsYaml = coilsYaml + "  c_trough_eject:" & vbCrLf
		coilsYaml = coilsYaml + "    number: " & coilsNumber & vbCrLf 
		coilsNumber = coilsNumber + 1

		For Each device in glf_ball_devices.Items()
			ballDevicesYaml = ballDevicesYaml + device.ToYaml()
			coilsYaml = coilsYaml + "  c_" & device.Name & "_eject:" & vbCrLf
			coilsYaml = coilsYaml + "    number: " & coilsNumber & vbCrLf 
			coilsNumber = coilsNumber + 1
		Next
		For Each device in Glf_ShotProfiles.Items()
			shotProfilesYaml = shotProfilesYaml + device.ToYaml()
		Next

		Dim init_index
    	Dim init_var_keys : init_var_keys = glf_initialVars.Keys()
    	Dim init_var_items : init_var_items = glf_initialVars.Items()
    	For init_index=0 To UBound(init_var_keys)
			playerVarsYaml = playerVarsYaml + "  " & init_var_keys(init_index) & ":" & vbCrLf
			playerVarsYaml = playerVarsYaml + "    initial_value: " & init_var_items(init_index) & vbCrLf
			Select Case VarType(init_var_items(init_index))
				Case vbInteger, vbLong
					playerVarsYaml = playerVarsYaml + "    value_type: int" & vbCrLf
				Case vbSingle, vbDouble
					playerVarsYaml = playerVarsYaml + "    value_type: float" & vbCrLf
				Case vbString
					playerVarsYaml = playerVarsYaml + "    value_type: str" & vbCrLf
        	End Select
    	Next

		Dim fso, modesFolder, TxtFileStream, monitorFolder, configFolder, showsFolder
		Set fso = CreateObject("Scripting.FileSystemObject")
		monitorFolder = "glf_mpf\monitor\"
		showsFolder = "glf_mpf\shows\"
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
		If Not fso.FolderExists("glf_mpf\shows") Then
			fso.CreateFolder "glf_mpf\shows"
		End If
		Set TxtFileStream = fso.OpenTextFile(monitorFolder & "\monitor.yaml", 2, True)
		TxtFileStream.WriteLine monitorYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\config.yaml", 2, True)
		TxtFileStream.WriteLine configYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\ball_devices.yaml", 2, True)
		TxtFileStream.WriteLine ballDevicesYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\coils.yaml", 2, True)
		TxtFileStream.WriteLine coilsYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\shot_profiles.yaml", 2, True)
		TxtFileStream.WriteLine shotProfilesYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\player_vars.yaml", 2, True)
		TxtFileStream.WriteLine playerVarsYaml
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
		Dim showsYaml
		For Each device in glf_shows.Items()
			If Not device.Name = "flash" And Not device.Name = "flash_color" And Not device.Name = "led_color" And Not device.Name = "off" And Not device.Name = "on" Then
				showsYaml = "#show_version=6" & vbCrLf & vbCrLf
				showsYaml = showsYaml + device.ToYaml()
				Set TxtFileStream = fso.OpenTextFile(showsFolder & "\" & device.Name & ".yaml", 2, True)
				TxtFileStream.WriteLine showsYaml
				TxtFileStream.Close
			End If
		Next
		
		For Each device in glf_modes.Items()
			device.ToYaml()
		Next

		glf_debugLog.WriteToLog "Init", "Finished MPF Config"

		'***GLFMPF_EXPORT_END***
	End If

	'Cache Shows
	glf_debugLog.WriteToLog "Init", "Caching Shows"
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
						Dim state_tokens : Set state_tokens = state.Tokens()
						For Each key In state_tokens.Keys()
							mergedTokens.Add key, state_tokens(key)
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
				Dim state_items : state_items = mode_state_machine.StateItems
				For x=0 to UBound(state_items)
					Set state = state_items(x)
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
	glf_debugLog.WriteToLog "Init", "Finished Caching Shows"

	glf_debugLog.WriteToLog "Init", "Creating Machine Vars"
	With CreateMachineVar("player1_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With
	With CreateMachineVar("player2_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With
	With CreateMachineVar("player3_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With
	With CreateMachineVar("player4_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With

	If Not IsNull(glf_highscore) Then
		glf_highscore.WriteDefaults()
	End If

	Glf_ReadMachineVars("MachineVars")
	Glf_ReadMachineVars("HighScores")
	glf_debugLog.WriteToLog "Init", "Finished Creating Machine Vars"
	'glf_debugLog.WriteToLog "Code String", glf_codestr
	If glf_production_mode = False Then
		Dim fso1, TxtFileStream1
		Set fso1 = CreateObject("Scripting.FileSystemObject")
		Set TxtFileStream1 = fso1.OpenTextFile("cached-functions.vbs", 2, True)
		TxtFileStream1.WriteLine glf_codestr
		TxtFileStream1.Close
	End If

	SetDelay "reset", "Glf_Reset", Null, 1000
End Sub

Sub Glf_Reset(args)
	DispatchQueuePinEvent "reset_complete", Null
End Sub

AddPinEventListener "reset_complete", "initial_segment_displays", "Glf_SegmentInit", 100, Null
AddPinEventListener "reset_virtual_segment_lights", "reset_segment_displays", "Glf_SegmentInit", 100, Null
Function Glf_SegmentInit(args)
	Dim segment_display
	For Each segment_display in glf_segment_displays.Items()	
		segment_display.SetVirtualDMDLights Not glf_flex_alphadmd_enabled
	Next
	If Not IsNull(args) Then
		If IsObject(args(1)) Then
			Set Glf_SegmentInit = args(1)
		Else
			Glf_SegmentInit = args(1)
		End If
	Else
		Glf_SegmentInit = Null
	End If
End Function

Sub Glf_ReadMachineVars(section)
    Dim objFSO, objFile, arrLines, line, inSection
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not objFSO.FileExists(CGameName & "_glf.ini") Then Exit Sub
    
    Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
    arrLines = Split(objFile.ReadAll, vbCrLf)
    objFile.Close
    
    inSection = False
    For Each line In arrLines
        line = Trim(line)
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase(section))
        ElseIf inSection And InStr(line, "=") > 0 Then
			Dim split_key : split_key = Split(line, "=")
			Dim key : key = Trim(split_key(0))
			If glf_machine_vars.Exists(key) Then
	            glf_machine_vars(key).Value = Trim(split_key(1))
			Else	
				With CreateMachineVar(key)
					.InitialValue = Trim(split_key(1))
					.ValueType = "int"
					.Persist = False
				End With
			End If
        End If
    Next
End Sub

Sub Glf_DisableVirtualSegmentDmd()
	If Not glf_flex_alphadmd is Nothing Then
		glf_flex_alphadmd.Show = False
		glf_flex_alphadmd.Run = False
		Set glf_flex_alphadmd = Nothing
	End If
	glf_flex_alphadmd_enabled = False
	DispatchPinEvent "reset_virtual_segment_lights", Null
End Sub

Sub Glf_EnableVirtualSegmentDmd(args)
	Glf_DisableVirtualSegmentDmd()
	Dim i
	Set glf_flex_alphadmd = CreateObject("FlexDMD.FlexDMD")
	With glf_flex_alphadmd
		.TableFile = Table1.Filename & ".vpx"
		.Color = RGB(255, 88, 32)
		.Width = 128
		.Height = 32
		.Clear = True
		.Run = True
		.GameName = cGameName
		.RenderMode = 3
	End With
	For i = 0 To 31
		glf_flex_alphadmd_segments(i) = 0
	Next
	glf_flex_alphadmd.Segments = glf_flex_alphadmd_segments
	glf_flex_alphadmd_enabled = True
	DispatchPinEvent "reset_virtual_segment_lights", Null
End Sub

Sub Glf_WriteMachineVars(section)
    Dim objFSO, objFile, arrLines, line, inSection, foundSection
    Dim outputLines, key
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(CGameName & "_glf.ini") Then
        Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
        arrLines = Split(objFile.ReadAll, vbCrLf)
        objFile.Close
    Else
        arrLines = Array()
    End If
    
    outputLines = ""
    inSection = False
    foundSection = False
    
    For Each line In arrLines
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase(section))
            foundSection = foundSection Or inSection
        End If
        
        If inSection And InStr(line, "=") > 0 Then
			Dim split_key : split_key = Split(line, "=")
            key = Trim(split_key(0))
            If glf_machine_vars.Exists(key) Then
				If glf_machine_vars(key).Persist = True Then
                	line = key & "=" & glf_machine_vars(key).Value
                	glf_machine_vars.Remove key
				End If
            End If
        End If

		If line = "" And inSection Then
			' Add remaining keys in the section
			For Each key In glf_machine_vars.Keys
				If glf_machine_vars(key).Persist = True Then
					outputLines = outputLines & key & "=" & glf_machine_vars(key).Value & vbCrLf
				End If
			Next
			glf_machine_vars.RemoveAll
		End If
        If line <> "" Then
			outputLines = outputLines & line & vbCrLf
		End If
    Next
    
    If Not foundSection Then
        outputLines = outputLines & "["&section&"]" & vbCrLf
        For Each key In glf_machine_vars.Keys
			If glf_machine_vars(key).Persist = True Then
            	outputLines = outputLines & key & "=" & glf_machine_vars(key).Value & vbCrLf
			End If
        Next
    End If
    
    Set objFile = objFSO.CreateTextFile(CGameName & "_glf.ini", True)
    objFile.Write outputLines
    objFile.Close
End Sub



Sub Glf_Options(ByVal eventId)
	
	Dim glfMaxDispatch : glfMaxDispatch = 1

	'***GLF_DEBUG_OPTIONS_START***
	Dim glfDebug : glfDebug = Table1.Option("Glf Debug Log", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfDebug = 1 Then
		glf_debugEnabled = True
		glf_debugLog.EnableLogs
	Else
		glf_debugEnabled = False
		glf_debugLog.DisableLogs
	End If
	glf_debugLog.WriteToLog "Options", "Start"

	Dim glfDebugLevel : glfDebugLevel = Table1.Option("Glf Debug Log Level", 0, 1, 1, 0, 0, Array("Info", "Debug"))
	If glfDebugLevel = 1 Then
		glf_debug_level = "Debug"
	Else
		glf_debug_level = "Info"
	End If

	glfMaxDispatch = Table1.Option("Glf Frame Dispatch", 1, 10, 1, 1, 0, Array("5", "10", "15", "20", "25", "30", "35", "40", "45", "50"))

	glf_debugLog.WriteToLog "Options", "GLF Monitor Check"
	Dim glfuseDebugBCP : glfuseDebugBCP = Table1.Option("Glf Monitor", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseDebugBCP = 1 And useGlfBCPMonitor = False Then
		useGlfBCPMonitor = True
		If glf_hasDebugController = False Then
			SetDelay "start_glf_monitor", "Glf_ConnectToDebugBCPMediaController", Null, 500
		End If
	ElseIf glfuseDebugBCP = 0 And useGlfBCPMonitor = True Then
		useGlfBCPMonitor = False
		If Not IsNull(glf_debugBcpController) Then
			glf_debugBcpController.Disconnect
			glf_debugBcpController = Null
			glf_hasDebugController = False
		End If
	End If
	'***GLF_DEBUG_OPTIONS_END***
	Dim ballsPerGame : ballsPerGame = Table1.Option("Balls Per Game", 1, 2, 1, 1, 0, Array("3 Balls", "5 Balls"))
	If ballsPerGame = 1 Then
		glf_game.BallsPerGame = 3
	Else
		glf_game.BallsPerGame = 5
	End If
	Dim tilt_sensitivity : tilt_sensitivity = Table1.Option("Tilt Sensitivity (digital nudge)", 1, 10, 1, 5, 0, Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10"))
	glf_tilt_sensitivity = tilt_sensitivity

	glf_max_dispatch = glfMaxDispatch*5

	glf_debugLog.WriteToLog "Options", "BCP Check"
	Dim glfuseBCP : glfuseBCP = Table1.Option("Glf Backbox Control Protocol", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseBCP = 1 Then
		If IsNull(bcpController) Then
			SetDelay "start_glf_bcp", "Glf_ConnectToBCPMediaController", Null, 500
		End If
	Else
		useBCP = False
		If Not IsNull(bcpController) Then
			bcpController.Disconnect
			bcpController = Null
		End If
	End If

	glf_debugLog.WriteToLog "Options", "GLF Segments (Flex) Check"
	Dim glfuseVirtualSegmentDMD : glfuseVirtualSegmentDMD = Table1.Option("Glf Virtual Segment DMD", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseVirtualSegmentDMD = 1 And glf_flex_alphadmd_enabled = False Then
		SetDelay "start_flex_segments", "Glf_EnableVirtualSegmentDmd", Null, 500
	ElseIf glfuseVirtualSegmentDMD = 0 And  glf_flex_alphadmd_enabled = True Then
		Glf_DisableVirtualSegmentDmd()
	End If

	glf_debugLog.WriteToLog "Options", "LightmapSync"
	Dim min_lightmap_update_rate : min_lightmap_update_rate = Table1.Option("Glf Min Lightmap Update Rate", 1, 6, 1, 1, 0, Array("Disabled", "30 Hz", "60 Hz", "120 Hz", "144 Hz", "165 Hz"))
    Select Case min_lightmap_update_rate
		Case 1: glf_max_lightmap_sync_enabled = False
		Case 2: glf_max_lightmap_sync = 30 : glf_max_lightmap_sync_enabled = True
        Case 3: glf_max_lightmap_sync = 15 : glf_max_lightmap_sync_enabled = True
        Case 4: glf_max_lightmap_sync = 8 : glf_max_lightmap_sync_enabled = True
        Case 5: glf_max_lightmap_sync = 7 : glf_max_lightmap_sync_enabled = True
        Case 6: glf_max_lightmap_sync = 6 : glf_max_lightmap_sync_enabled = True
    End Select
	
End Sub

Public Sub Glf_Exit()
	If Not IsNull(bcpController) Then
		bcpController.Disconnect
		bcpController = Null
	End If
	If Not IsNull(glf_debugBcpController) Then
		glf_debugBcpController.Disconnect
		glf_debugBcpController = Null
		glf_hasDebugController = False
	End If
	If glf_debugEnabled = True Then
		glf_debugLog.WriteToLog "Max Lights", glf_max_lights_test
		glf_debugLog.DisableLogs
	End If
	Glf_DisableVirtualSegmentDmd()
	Glf_WriteMachineVars("MachineVars")
End Sub

Public Sub Glf_KeyDown(ByVal keycode)
    If glf_gameStarted = True Then
		If keycode = StartGameKey Then
			If glf_canAddPlayers = True Then
				Glf_AddPlayer()
			End If
			DispatchPinEvent "s_start_active", True
		End If
	Else
		If keycode = StartGameKey Then
			DispatchPinEvent "s_start_active", True
		End If
	End If

	If keycode = LeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_flipper_active", Null
	End If
	
	If keycode = RightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_flipper_active", Null
	End If
	
	If keycode = LockbarKey Then
		RunAutoFireDispatchPinEvent "s_lockbar_key_active", Null
	End If

	If KeyCode = PlungerKey Then
		RunAutoFireDispatchPinEvent "s_plunger_key_active", Null
	End If

	If KeyCode = LeftMagnaSave Then
		RunAutoFireDispatchPinEvent "s_left_magna_key_active", Null
	End If

	If KeyCode = RightMagnaSave Then
		RunAutoFireDispatchPinEvent "s_right_magna_key_active", Null
	End If

	If KeyCode = StagedRightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_staged_flipper_key_active", Null
	End If

	If KeyCode = StagedLeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_staged_flipper_key_active", Null
	End If
	
	If keycode = MechanicalTilt Then 
		SetDelay "glf_mechcanical_tilt_debounce", "MechcanicalTiltDebounce", Null, 300
    End If

	If keycode = LeftTiltKey Then 
		Nudge 90, 2
		Glf_CheckTilt
	End If
    If keycode = RightTiltKey Then
		Nudge 270, 2
		Glf_CheckTilt
	End If
    If keycode = CenterTiltKey Then 
		Nudge 0, 3
		Glf_CheckTilt
	End If

	If KeyCode = AddCreditKey Then
		RunAutoFireDispatchPinEvent "s_add_credit_key_active", Null
	End If

	If KeyCode = AddCreditKey2 Then
		RunAutoFireDispatchPinEvent "s_add_credit_key2_active", Null
	End If
End Sub

Public Sub Glf_KeyUp(ByVal keycode)
	

	If keycode = StartGameKey Then
		DispatchRelayPinEvent "request_to_start_game", True
		DispatchPinEvent "s_start_inactive", True
	End If

	If KeyCode = PlungerKey Then
		RunAutoFireDispatchPinEvent "s_plunger_key_inactive", Null
	End If

	If keycode = LeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_flipper_inactive", Null
	End If
	
	If keycode = RightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_flipper_inactive", Null
	End If

	If keycode = LockbarKey Then
		RunAutoFireDispatchPinEvent "s_lockbar_key_inactive", Null
	End If

	If KeyCode = LeftMagnaSave Then
		RunAutoFireDispatchPinEvent "s_left_magna_key_inactive", Null
	End If

	If KeyCode = RightMagnaSave Then
		RunAutoFireDispatchPinEvent "s_right_magna_key_inactive", Null
	End If

	If KeyCode = StagedRightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_staged_flipper_key_inactive", Null
	End If

	If KeyCode = StagedLeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_staged_flipper_key_inactive", Null
	End If		

	If KeyCode = AddCreditKey Then
		RunAutoFireDispatchPinEvent "s_add_credit_key_inactive", Null
	End If

	If KeyCode = AddCreditKey2 Then
		RunAutoFireDispatchPinEvent "s_add_credit_key2_inactive", Null
	End If
End Sub

Public Sub MechcanicalTiltDebounce(args)
	RunAutoFireDispatchPinEvent "s_tilt_warning_active", Null
End Sub

Dim glf_lastEventExecutionTime, glf_lastBcpExecutionTime, glf_lastLightUpdateExecutionTime, glf_lastTiltUpdateExecutionTime
glf_lastEventExecutionTime = 0
glf_lastBcpExecutionTime = 0
glf_lastLightUpdateExecutionTime = 0
glf_lastTiltUpdateExecutionTime = 0

Public Sub Glf_GameTimer_Timer()

	If (gametime - glf_lastEventExecutionTime) > 25 Then
		'debug.print "Slow GLF Frame: " & gametime - glf_lastEventExecutionTime & ". Dispatch Count: " & glf_frame_dispatch_count & ". Handler Count: " & glf_frame_handler_count
	End If
	glf_frame_dispatch_count = 0
	glf_frame_handler_count = 0
	'glf_temp1 = 0

	Dim i, key, keys, lightMap
	i = 0
	keys = glf_dispatch_handlers_await.Keys()
	i = Glf_RunHandlers(i)
	If i<glf_max_dispatch Then
		keys = glf_dispatch_await.Keys()
		For Each key in keys
			RunDispatchPinEvent key, glf_dispatch_await(key)
			glf_dispatch_await.Remove key
			If UBound(glf_dispatch_handlers_await.Keys())>-1 Then
				'Handlers were added, process those first.
				i = Glf_RunHandlers(i)
			End If
			i = i + 1
			If i>=glf_max_dispatch Then
				Exit For
			End If
		Next
	End If

	DelayTick

	If glf_max_lightmap_sync_enabled = True Then
		keys = glf_dispatch_lightmaps_await.Keys()
		'debug.print(ubound(keys))
		If glf_max_lights_test < Ubound(keys) Then
			glf_max_lights_test = Ubound(keys)
		End If
		For Each key in keys
			For Each lightMap in glf_lightMaps(key)
				lightMap.Color = glf_lightNames(key).Color
			Next
			glf_dispatch_lightmaps_await.Remove key
			If (gametime - glf_lastEventExecutionTime) > glf_max_lightmap_sync Then
				'debug.print("Exiting")
				Exit For
			End If
		Next
	End If
	If (gametime - glf_lastTiltUpdateExecutionTime) >=50 And glf_current_virtual_tilt > 0 Then
		glf_current_virtual_tilt = glf_current_virtual_tilt - 0.1
		glf_lastTiltUpdateExecutionTime = gametime
		'Debug.print("Tilt Cooldown: " & glf_current_virtual_tilt) 
	End If

	If (gametime - glf_lastBcpExecutionTime) >= 300 Then
        glf_lastBcpExecutionTime = gametime
		Glf_BcpUpdate
		Glf_MonitorPlayerStateUpdate "GLF BIP", glf_BIP
		Glf_MonitorBcpUpdate
    End If

	If glf_last_switch_hit_time > 0 And (gametime - glf_last_ballsearch_reset_time) > 2000 Then
		Glf_ResetBallSearch
		glf_last_switch_hit_time = 0
		glf_last_ballsearch_reset_time = gametime
	End If
	glf_lastEventExecutionTime = gametime
End Sub

Sub Glf_CheckTilt()
	glf_current_virtual_tilt = glf_current_virtual_tilt + glf_tilt_sensitivity
	If (glf_current_virtual_tilt > 10) Then 
		RunAutoFireDispatchPinEvent "s_tilt_warning_active", Null
		glf_current_virtual_tilt = glf_tilt_sensitivity
	End If
End Sub

Sub Glf_ResetBallSearch()
	If glf_ballsearch_enabled = True Then
		glf_ballsearch.Reset()
	End If
End Sub

Public Function Glf_RunHandlers(i)
	Dim key, keys, args
	keys = glf_dispatch_handlers_await.Keys()
	If UBound(keys) = -1 Then
		Glf_RunHandlers = i
		Exit Function
	End If
	For Each key in keys
		args = glf_dispatch_handlers_await(key)
		Dim wait_for : wait_for = DispatchPinHandlers(key, args)
		glf_dispatch_handlers_await.Remove key
		If Not IsEmpty(wait_for) Then
			Dim remaining_handlers_keys : remaining_handlers_keys = glf_dispatch_handlers_await.Keys
			Dim remaining_handlers_items : remaining_handlers_items = glf_dispatch_handlers_await.Items
			AddPinEventListener wait_for, key & "_wait_for", "ContinueDispatchQueuePinEvent", 1000, Array(remaining_handlers_keys, remaining_handlers_items)
			glf_dispatch_handlers_await.RemoveAll
			Exit For
		End If
		i = i + 1
		If i=glf_max_dispatch Then
			Exit For
		End If
	Next
	If Ubound(glf_dispatch_handlers_await.Keys())=-1 Then
		'Finished processing Handlers for current event.
		'Remove any blocks for this event.
		Glf_EventBlocks(args(2)).RemoveAll
	End If
	Glf_RunHandlers = i
End Function
Dim glf_tmp_lmarr
Public Function Glf_RegisterLights()

	If glf_production_mode = False Then
		Dim elementDict : Set elementDict = CreateObject("Scripting.Dictionary")

		For Each e in GetElements()
			If typename(e) = "Primitive" or typename(e) = "Flasher"  Then
				elementDict.Add LCase(e.Name), True
			End If
		Next
	End If

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
		If glf_production_mode = False Then
			Dim e, lmStr: lmStr = "Dim glf_" & light.name & "_lmarr : glf_" & light.name & "_lmarr = Array("    
			For Each e in elementDict.Keys
				If InStr(e, LCase("_" & light.Name & "_")) Then
					lmStr = lmStr & e & ","
				End If
				For Each tag in tags
					tag = "T_" & Trim(tag)
					If InStr(e, LCase("_" & tag & "_")) Then
						lmStr = lmStr & e & ","
					End If
				Next
			Next
			lmStr = lmStr & "Null)"
			lmStr = Replace(lmStr, ",Null)", ")")
			lmStr = Replace(lmStr, "Null)", ")")
			ExecuteGlobal lmStr
			glf_lightMaps.Add light.Name, Eval("glf_" & light.name & "_lmarr")
			glf_codestr = glf_codestr & lmStr & vbCrLf
			glf_codestr = glf_codestr & "glf_lightMaps.Add """ & light.name & """, glf_" & light.Name & "_lmarr" & vbCrLf
		End If

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
	
	glf_lightNames(light).Color = rgbColor

	If glf_max_lightmap_sync_enabled = True Then
		If Not glf_dispatch_lightmaps_await.Exists(light) Then
			glf_dispatch_lightmaps_await.Add light, True
		End If
	Else
		dim lightMap
		For Each lightMap in glf_lightMaps(light)
			lightMap.Color = glf_lightNames(light).Color
		Next
	End If
End Function

Public Function Glf_ParseInput(value)
	Dim templateCode : templateCode = ""
	Dim tmp: tmp = value
	Dim isVariable, parts
	If glf_funcRefMap.Exists(CStr(value)) Then
		'msgbox "match " & value & " REF: " & glf_funcRefMap(value)
		Glf_ParseInput = Array(glf_funcRefMap(CStr(value)), value, True)
	Else
		Select Case VarType(value)
			Case 8 ' vbString
				tmp = Glf_ReplaceCurrentPlayerAttributes(tmp)
				tmp = Glf_ReplaceAnyPlayerAttributes(tmp)
				tmp = Glf_ReplaceDeviceAttributes(tmp)
				tmp = Glf_ReplaceMachineAttributes(tmp)
				tmp = Glf_ReplaceModeAttributes(tmp)
				tmp = Glf_ReplaceGameAttributes(tmp)
				tmp = Glf_ReplaceKwargsAttributes(tmp)
				'msgbox tmp
				If InStr(tmp, " if ") Then
					templateCode = "Function Glf_" & glf_FuncCount & "(args)" & vbCrLf
					templateCode = templateCode & vbTab & Glf_ConvertIf(tmp, "Glf_" & glf_FuncCount) & vbCrLf
					templateCode = templateCode & "End Function"
				Else
					isVariable = Glf_IsCondition(tmp)
					If Not IsNull(isVariable) Then
						'The input needs formatting
						parts = Split(isVariable, ":")
						If UBound(parts) = 1 Then
							tmp = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
						Else
							tmp = parts(0)
						End If
					End If
					templateCode = "Function Glf_" & glf_FuncCount & "(args)" & vbCrLf
					templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
					templateCode = templateCode & "End Function"
				End IF
			Case Else
				templateCode = "Function Glf_" & glf_FuncCount & "(args)" & vbCrLf			
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
		glf_codestr = glf_codestr & templateCode & vbCrLf
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		If Not glf_funcRefMap.Exists(CStr(value)) Then
			glf_codestr = glf_codestr & "glf_funcRefMap.Add """ & Replace(value, """", """""") & """, """ & funcRef & """" & vbCrLf
			glf_funcRefMap.Add CStr(value), funcRef
		End If
		glf_FuncCount = glf_FuncCount + 1
		Glf_ParseInput = Array(funcRef, value, True)
	End If
End Function

Public Function Glf_ParseEventInput(value)
	Dim templateCode : templateCode = ""
	Dim parts : parts = Split(value, ":")
	Dim event_delay : event_delay = 0
	Dim priority : priority = 0
	If UBound(parts) = 1 Then
		value = parts(0)
		event_delay= parts(1)
	End If

	Dim condition : condition = Glf_IsCondition(value)
	If IsNull(condition) Then
		parts = Split(value, ".")
		If UBound(parts) = 1 Then
			value = parts(0)
			priority= parts(1)
		End If
		Glf_ParseEventInput = Array(value, value, Null, event_delay, priority)
	Else

		If glf_funcRefMap.Exists(value) Then
			Dim func_ref : func_ref = glf_funcRefMap(value)
			value = Replace(value, "{"&condition&"}", "")
			parts = Split(value, ".")
			If UBound(parts) = 1 Then
				value = parts(0)
				priority= parts(1)
			End If
			Glf_ParseEventInput = Array(value & func_ref, value, func_ref, event_delay, priority)
			Exit Function
		End If

		dim conditionReplaced : conditionReplaced = Glf_ReplaceCurrentPlayerAttributes(condition)
		conditionReplaced = Glf_ReplaceAnyPlayerAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceDeviceAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceMachineAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceModeAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceGameAttributes(conditionReplaced)

		conditionReplaced = Glf_ReplaceKwargsAttributes(conditionReplaced)
		templateCode = "Function Glf_" & glf_FuncCount & "(args)" & vbCrLf
		templateCode = templateCode & vbTab & "On Error Resume Next" & vbCrLf
		templateCode = templateCode & vbTab & Glf_ConvertCondition(conditionReplaced, "Glf_" & glf_FuncCount) & vbCrLf
		templateCode = templateCode & vbTab & "If Err Then Glf_" & glf_FuncCount & " = False" & vbCrLf
		templateCode = templateCode & "End Function"
		ExecuteGlobal templateCode
		glf_codestr = glf_codestr & templateCode & vbCrLf
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		glf_FuncCount = glf_FuncCount + 1
		If Not glf_funcRefMap.Exists(value) Then
			glf_codestr = glf_codestr & "glf_funcRefMap.Add """ & Replace(value, """", """""") & """, """ & funcRef & """" & vbCrLf
			glf_funcRefMap.Add value, funcRef
		End If
		value = Replace(value, "{"&condition&"}", "")
		parts = Split(value, ".")
		If UBound(parts) = 1 Then
			value = parts(0)
			priority= parts(1)
		End If
		Glf_ParseEventInput = Array(value & funcRef ,value, funcRef, event_delay, priority)
	End If
End Function

Public Function Glf_ParseDispatchEventInput(value)
	Dim templateCode : templateCode = ""
	Dim kwargs : kwargs = Empty
	Dim eventKey
	Dim pos : pos = InStr(value, ":")
	If pos > 0 Then
		eventKey = Trim(Left(value, pos - 1))
		kwargs = Glf_IsCondition(Trim(Mid(value, pos + 1)))
	Else
		eventKey = value
	End If

	If IsEmpty(kwargs) Then
		Glf_ParseDispatchEventInput = Array(eventKey, Empty)
	Else
		dim kwargsReplaced : kwargsReplaced = Glf_ReplaceCurrentPlayerAttributes(kwargs)
		kwargsReplaced = Glf_ReplaceAnyPlayerAttributes(kwargsReplaced)
		kwargsReplaced = Glf_ReplaceDeviceAttributes(kwargsReplaced)
		kwargsReplaced = Glf_ReplaceMachineAttributes(kwargsReplaced)
		kwargsReplaced = Glf_ReplaceModeAttributes(kwargsReplaced)
		kwargsReplaced = Glf_ReplaceGameAttributes(kwargsReplaced)

		templateCode = "Function Glf_" & glf_FuncCount & "(args)" & vbCrLf
		templateCode = templateCode & vbTab & "On Error Resume Next" & vbCrLf
		templateCode = templateCode & vbTab & Glf_ConvertDynamicKwargs(kwargsReplaced, "Glf_" & glf_FuncCount) & vbCrLf
		templateCode = templateCode & vbTab & "If Err Then Glf_" & glf_FuncCount & " = Null" & vbCrLf
		templateCode = templateCode & "End Function"
		'msgbox templateCode
		ExecuteGlobal templateCode
		'glf_codestr = glf_codestr & templateCode & vbCrLf
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		glf_FuncCount = glf_FuncCount + 1

		Glf_ParseDispatchEventInput = Array(eventKey, funcRef)
	End If
End Function

' Function Glf_ReplaceCurrentPlayerAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "current_player\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
'     replacement = "GetPlayerState(""$1"")"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceCurrentPlayerAttributes = outputString
' End Function

Function Glf_ReplaceCurrentPlayerAttributes(inputString)
    Dim outputString, startPos, attrStart, attrEnd
    Dim beforeMatch, afterMatch, attribute, replacement

    outputString = inputString
    startPos = InStr(outputString, "current_player.")

    Do While startPos > 0
        ' Start of the attribute is just after the dot
        attrStart = startPos + Len("current_player.")
        attrEnd = attrStart

        ' Find the end of the attribute (until a non-word character or end of string)
        Do While attrEnd <= Len(outputString)
            Dim ch
            ch = Mid(outputString, attrEnd, 1)
            If Not (ch >= "a" And ch <= "z") And Not (ch >= "A" And ch <= "Z") And Not (ch >= "0" And ch <= "9") And ch <> "_" Then
                Exit Do
            End If
            attrEnd = attrEnd + 1
        Loop

        attribute = Mid(outputString, attrStart, attrEnd - attrStart)
        replacement = "GetPlayerState(""" & attribute & """)"

        beforeMatch = Left(outputString, startPos - 1)
        afterMatch = Mid(outputString, attrEnd)
        outputString = beforeMatch & replacement & afterMatch

        ' Search again from just after the replacement
        startPos = InStr(startPos + Len(replacement), outputString, "current_player.")
    Loop

    Glf_ReplaceCurrentPlayerAttributes = outputString
End Function


' Function Glf_ReplaceAnyPlayerAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "players\[([0-3]+)\]\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
'     replacement = "GetPlayerStateForPlayer($1, ""$2"")"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceAnyPlayerAttributes = outputString
' End Function

Function Glf_ReplaceAnyPlayerAttributes(inputString)
    Dim outputString, startPos, openBracket, closeBracket, dotPos
    Dim beforeMatch, afterMatch, indexStr, attribute, replacement

    outputString = inputString
    startPos = InStr(outputString, "players[")

    Do While startPos > 0
        openBracket = InStr(startPos, outputString, "[")
        closeBracket = InStr(openBracket, outputString, "]")
        dotPos = InStr(closeBracket + 1, outputString, ".")

        If openBracket > 0 And closeBracket > openBracket And dotPos > closeBracket Then
            indexStr = Mid(outputString, openBracket + 1, closeBracket - openBracket - 1)

            If IsNumeric(indexStr) Then
                If CInt(indexStr) >= 0 And CInt(indexStr) <= 3 Then
                    ' Extract attribute
                    Dim attrStart, attrEnd, ch
                    attrStart = dotPos + 1
                    attrEnd = attrStart
                    Do While attrEnd <= Len(outputString)
                        ch = Mid(outputString, attrEnd, 1)
                        If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                            Exit Do
                        End If
                        attrEnd = attrEnd + 1
                    Loop
                    attribute = Mid(outputString, attrStart, attrEnd - attrStart)

                    replacement = "GetPlayerStateForPlayer(" & indexStr & ", """ & attribute & """)"
                    beforeMatch = Left(outputString, startPos - 1)
                    afterMatch = Mid(outputString, attrEnd)
                    outputString = beforeMatch & replacement & afterMatch

                    startPos = InStr(startPos + Len(replacement), outputString, "players[")
                Else
                    ' Invalid index, move past it
                    startPos = InStr(closeBracket + 1, outputString, "players[")
                End If
            Else
                ' Not numeric index
                startPos = InStr(closeBracket + 1, outputString, "players[")
            End If
        Else
            Exit Do
        End If
    Loop

    Glf_ReplaceAnyPlayerAttributes = outputString
End Function


' Function Glf_ReplaceDeviceAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "device\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
' 	replacement = "glf_$1(""$2"").GetValue(""$3"")"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceDeviceAttributes = outputString
' End Function

Function Glf_ReplaceDeviceAttributes(inputString)
    Dim outputString, startPos, endPos, beforeMatch, afterMatch
    Dim fullMatch, parts, deviceType, deviceId, attribute

    outputString = inputString
    startPos = InStr(outputString, "device.")

    Do While startPos > 0
        ' Find the end of the match by looking for three dots after "device."
        Dim firstDot, secondDot, thirdDot

        firstDot = InStr(startPos + 8, outputString, ".")
        If firstDot = 0 Then Exit Do

        secondDot = InStr(firstDot + 1, outputString, ".")
        If secondDot = 0 Then Exit Do

        ' Third part ends at the next non-word character or end of string
        thirdDot = secondDot + 1
        Do While thirdDot <= Len(outputString)
            Dim ch
            ch = Mid(outputString, thirdDot, 1)
            If Not (ch >= "a" And ch <= "z") And Not (ch >= "A" And ch <= "Z") And Not (ch >= "0" And ch <= "9") And ch <> "_" Then
                Exit Do
            End If
            thirdDot = thirdDot + 1
        Loop
        endPos = thirdDot - 1

        ' Extract full match and parts
        fullMatch = Mid(outputString, startPos, endPos - startPos + 1)
        parts = Split(fullMatch, ".")
        If UBound(parts) = 3 Then
            deviceType = parts(1)
            deviceId = parts(2)
            attribute = parts(3)

            ' Build replacement
            Dim replacement
            replacement = "glf_" & deviceType & "(""" & deviceId & """).GetValue(""" & attribute & """)"

            ' Replace in outputString
            beforeMatch = Left(outputString, startPos - 1)
            afterMatch = Mid(outputString, endPos + 1)
            outputString = beforeMatch & replacement & afterMatch

            ' Move startPos forward
            startPos = InStr(startPos + Len(replacement), outputString, "device.")
        Else
            Exit Do
        End If
    Loop

    Glf_ReplaceDeviceAttributes = outputString
End Function


' Function Glf_ReplaceMachineAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "machine\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
' 	replacement = "glf_machine_vars(""$1"").GetValue()"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceMachineAttributes = outputString
' End Function

Function Glf_ReplaceMachineAttributes(inputString)
    Dim outputString, startPos, attrStart, attrEnd
    Dim beforeMatch, afterMatch, attribute, replacement

    outputString = inputString
    startPos = InStr(outputString, "machine.")

    Do While startPos > 0
        ' Position of attribute name starts after "machine."
        attrStart = startPos + Len("machine.")
        attrEnd = attrStart

        ' Read characters in the attribute (letters, digits, or underscore)
        Do While attrEnd <= Len(outputString)
            Dim ch
            ch = Mid(outputString, attrEnd, 1)
            If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                Exit Do
            End If
            attrEnd = attrEnd + 1
        Loop

        ' Extract attribute name
        attribute = Mid(outputString, attrStart, attrEnd - attrStart)

        ' Build replacement text
        replacement = "glf_machine_vars(""" & attribute & """).GetValue()"

        ' Reconstruct string with replacement
        beforeMatch = Left(outputString, startPos - 1)
        afterMatch = Mid(outputString, attrEnd)
        outputString = beforeMatch & replacement & afterMatch

        ' Continue searching after the replacement
        startPos = InStr(startPos + Len(replacement), outputString, "machine.")
    Loop

    Glf_ReplaceMachineAttributes = outputString
End Function


' Function Glf_ReplaceGameAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "game\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
' 	replacement = "Glf_GameVariable(""$1"")"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceGameAttributes = outputString
' End Function

Function Glf_ReplaceGameAttributes(inputString)
    Dim outputString, startPos, attrStart, attrEnd
    Dim beforeMatch, afterMatch, attribute, replacement

    outputString = inputString
    startPos = InStr(outputString, "game.")

    Do While startPos > 0
        ' Position of attribute name starts after "game."
        attrStart = startPos + Len("game.")
        attrEnd = attrStart

        ' Walk forward while characters are valid for identifier
        Do While attrEnd <= Len(outputString)
            Dim ch
            ch = Mid(outputString, attrEnd, 1)
            If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                Exit Do
            End If
            attrEnd = attrEnd + 1
        Loop

        attribute = Mid(outputString, attrStart, attrEnd - attrStart)
        replacement = "Glf_GameVariable(""" & attribute & """)"

        beforeMatch = Left(outputString, startPos - 1)
        afterMatch = Mid(outputString, attrEnd)
        outputString = beforeMatch & replacement & afterMatch

        startPos = InStr(startPos + Len(replacement), outputString, "game.")
    Loop

    Glf_ReplaceGameAttributes = outputString
End Function


' Function Glf_ReplaceModeAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "modes\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
' 	replacement = "glf_modes(""$1"").GetValue(""$2"")"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceModeAttributes = outputString
' End Function

Function Glf_ReplaceModeAttributes(inputString)
    Dim outputString, startPos, modeStart, modeEnd, attrStart, attrEnd
    Dim beforeMatch, afterMatch, modeName, attribute, replacement

    outputString = inputString
    startPos = InStr(outputString, "modes.")

    Do While startPos > 0
        modeStart = startPos + Len("modes.")
        modeEnd = modeStart

        ' Read the mode name
        Do While modeEnd <= Len(outputString)
            Dim ch
            ch = Mid(outputString, modeEnd, 1)
            If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                Exit Do
            End If
            modeEnd = modeEnd + 1
        Loop

        ' Ensure there's a dot after mode name
        If Mid(outputString, modeEnd, 1) = "." Then
            ' Now read the attribute
            attrStart = modeEnd + 1
            attrEnd = attrStart

            Do While attrEnd <= Len(outputString)
                ch = Mid(outputString, attrEnd, 1)
                If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                    Exit Do
                End If
                attrEnd = attrEnd + 1
            Loop

            modeName = Mid(outputString, modeStart, modeEnd - modeStart)
            attribute = Mid(outputString, attrStart, attrEnd - attrStart)

            replacement = "glf_modes(""" & modeName & """).GetValue(""" & attribute & """)"

            beforeMatch = Left(outputString, startPos - 1)
            afterMatch = Mid(outputString, attrEnd)
            outputString = beforeMatch & replacement & afterMatch

            startPos = InStr(startPos + Len(replacement), outputString, "modes.")
        Else
            ' Not a valid modes.mode.attribute, skip ahead
            startPos = InStr(modeEnd + 1, outputString, "modes.")
        End If
    Loop

    Glf_ReplaceModeAttributes = outputString
End Function

' Function Glf_ReplaceKwargsAttributes(inputString)
'     Dim pattern, replacement, regex, outputString
'     pattern = "kwargs\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = True
'     replacement = "glf_dispatch_current_kwargs(""$1"")"
'     outputString = regex.Replace(inputString, replacement)
'     Set regex = Nothing
'     Glf_ReplaceKwargsAttributes = outputString
' End Function

Function Glf_ReplaceKwargsAttributes(inputString)
    Dim outputString, startPos, attrStart, attrEnd
    Dim beforeMatch, afterMatch, attribute, replacement

    outputString = inputString
    startPos = InStr(outputString, "kwargs.")

    Do While startPos > 0
        attrStart = startPos + Len("kwargs.")
        attrEnd = attrStart

        ' Walk forward through the attribute name
        Do While attrEnd <= Len(outputString)
            Dim ch
            ch = Mid(outputString, attrEnd, 1)
            If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                Exit Do
            End If
            attrEnd = attrEnd + 1
        Loop

        attribute = Mid(outputString, attrStart, attrEnd - attrStart)

        replacement = "glf_dispatch_current_kwargs(""" & attribute & """)"

        beforeMatch = Left(outputString, startPos - 1)
        afterMatch = Mid(outputString, attrEnd)
        outputString = beforeMatch & replacement & afterMatch

        startPos = InStr(startPos + Len(replacement), outputString, "kwargs.")
    Loop

    Glf_ReplaceKwargsAttributes = outputString
End Function


Function Glf_GameVariable(value)
	Glf_GameVariable = False
	Select Case value
		Case "tilted"
			Glf_GameVariable = glf_gameTilted
		Case "balls_per_game"
			Glf_GameVariable = glf_game.BallsPerGame()
		Case "balls_in_play"
			Glf_GameVariable = glf_BIP
	End Select
End Function

' Function Glf_CheckForGetPlayerState(inputString)
'     Dim pattern, regex, matches, match, hasGetPlayerState, attribute, playerNumber
' 	inputString = Glf_ReplaceCurrentPlayerAttributes(inputString)
' 	inputString = Glf_ReplaceAnyPlayerAttributes(inputString)
'     pattern = "GetPlayerState\(""([a-zA-Z0-9_]+)""\)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = False
    
'     Set matches = regex.Execute(inputString)
'     If matches.Count > 0 Then
'         hasGetPlayerState = True
' 		playerNumber = -1 'Current Player
'         attribute = matches(0).SubMatches(0)
'     Else
'         hasGetPlayerState = False
'         attribute = ""
' 		playerNumber = Null
'     End If


' 	pattern = "GetPlayerStateForPlayer\(([0-3]), ""([a-zA-Z0-9_]+)""\)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = False
    
'     Set matches = regex.Execute(inputString)
'     If matches.Count > 0 Then
'         hasGetPlayerState = True
' 		playerNumber = Int(matches(0).SubMatches(0))
'         attribute = matches(0).SubMatches(1)
'     Else
'         hasGetPlayerState = False
'         attribute = ""
' 		playerNumber = Null
'     End If

'     Set regex = Nothing
'     Set matches = Nothing
    
'     Glf_CheckForGetPlayerState = Array(hasGetPlayerState, attribute, playerNumber)
' End Function

Function Glf_CheckForGetPlayerState(inputString)
    Dim hasGetPlayerState, attribute, playerNumber
    Dim pos, startAttr, endAttr, ch, temp

    ' First, apply replacements
    inputString = Glf_ReplaceCurrentPlayerAttributes(inputString)
    inputString = Glf_ReplaceAnyPlayerAttributes(inputString)

    ' Check for GetPlayerState("...")
    pos = InStr(inputString, "GetPlayerState(""")
    If pos > 0 Then
        startAttr = pos + Len("GetPlayerState(""")
        endAttr = startAttr
        Do While endAttr <= Len(inputString)
            ch = Mid(inputString, endAttr, 1)
            If ch = """" Then Exit Do
            endAttr = endAttr + 1
        Loop

        attribute = Mid(inputString, startAttr, endAttr - startAttr)
        hasGetPlayerState = True
        playerNumber = -1 ' current player
    Else
        ' Check for GetPlayerStateForPlayer(n, "...")
        pos = InStr(inputString, "GetPlayerStateForPlayer(")
        If pos > 0 Then
            ' Extract player number
            startAttr = pos + Len("GetPlayerStateForPlayer(")
            endAttr = InStr(startAttr, inputString, ",")
            temp = Trim(Mid(inputString, startAttr, endAttr - startAttr))
            If IsNumeric(temp) Then
                playerNumber = CInt(temp)
            Else
                playerNumber = Null
            End If

            ' Extract attribute name
            startAttr = InStr(endAttr, inputString, """") + 1
            endAttr = InStr(startAttr, inputString, """")
            attribute = Mid(inputString, startAttr, endAttr - startAttr)

            hasGetPlayerState = True
        Else
            hasGetPlayerState = False
            attribute = ""
            playerNumber = Null
        End If
    End If

    Glf_CheckForGetPlayerState = Array(hasGetPlayerState, attribute, playerNumber)
End Function

' Function Glf_CheckForDeviceState(inputString)
'     Dim pattern, regex, matches, match, hasDeviceState, attribute, deviceType, deviceName
'     pattern = "devices\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
'     Set regex = New RegExp
'     regex.Pattern = pattern
'     regex.IgnoreCase = True
'     regex.Global = False
'     Set matches = regex.Execute(inputString)
'     If matches.Count > 0 Then
'         hasDeviceState = True
' 		deviceType = matches(0).SubMatches(0)
' 		deviceType = Left(deviceType, Len(deviceType)-1)
'         deviceName = matches(0).SubMatches(1)
' 		attribute = matches(0).SubMatches(2)
' 		attribute = Left(attribute, Len(attribute)-1)
'     Else
'         hasDeviceState = False
'         deviceType = Empty
' 		deviceName = Empty
' 		attribute = Empty
'     End If

'     Set regex = Nothing
'     Set matches = Nothing
    
'     Glf_CheckForDeviceState = Array(hasDeviceState, deviceType, deviceName, attribute)
' End Function

Function Glf_CheckForDeviceState(inputString)
    Dim hasDeviceState, deviceType, deviceName, attribute
    Dim pos, segStart, segEnd, part1, part2, part3
    Dim tempStr

    hasDeviceState = False
    deviceType = Empty
    deviceName = Empty
    attribute = Empty

    pos = InStr(inputString, "device.")
    If pos > 0 Then
        segStart = pos + Len("device.")
        
        ' Get first segment (deviceType)
        segEnd = InStr(segStart, inputString, ".")
        If segEnd > 0 Then
            part1 = Mid(inputString, segStart, segEnd - segStart)
            segStart = segEnd + 1

            ' Get second segment (deviceName)
            segEnd = InStr(segStart, inputString, ".")
            If segEnd > 0 Then
                part2 = Mid(inputString, segStart, segEnd - segStart)
                segStart = segEnd + 1

                ' Get third segment (attribute)
                segEnd = segStart
                Do While segEnd <= Len(inputString)
                    Dim ch
                    ch = Mid(inputString, segEnd, 1)
                    If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
                        Exit Do
                    End If
                    segEnd = segEnd + 1
                Loop
                part3 = Mid(inputString, segStart, segEnd - segStart)

                If part1 <> "" And part2 <> "" And part3 <> "" Then
                    hasDeviceState = True
                    deviceType = Left(part1, Len(part1) - 1) ' mimic your original trimming
                    deviceName = part2
                    attribute = Left(part3, Len(part3) - 1)  ' mimic your original trimming
                End If
            End If
        End If
    End If

    Glf_CheckForDeviceState = Array(hasDeviceState, deviceType, deviceName, attribute)
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
	value = Replace(value, "||", "Or")
	Glf_ConvertCondition = "    "&retName&" = " & value
End Function

Function Glf_ConvertDynamicKwargs(input, retName)
	'Value is a string of key values 
	Dim templateCode : templateCode = ""
	Dim arrayOfValues : arrayOfValues = Split(input, ",")
	Dim kvp
	templateCode = templateCode & vbTab & "Dim kwargs : Set kwargs = GlfKwargs()" & vbCrLf 
	For Each kvp in arrayOfValues
		Dim key,value
		Dim split_key : split_key = Split(kvp, ":")
		key = split_key(0)
		split_key = Split(kvp, ":")
		value = split_key(1)
		templateCode = templateCode & vbTab & "kwargs.Add """ & key & """, " & value & vbCrLf
	Next

	templateCode = templateCode & vbTab & "Set " & retName & " = kwargs"
	
	Glf_ConvertDynamicKwargs = templateCode
End Function

Function Glf_FormatValue(value, formatString)
    Dim padChar, width, result, align, hasCommas
	
	If TypeName(value) = "Boolean" Then
        If Not value Then
            Glf_FormatValue = ""
            Exit Function
        End If
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

Public Sub Glf_SetInitialPlayerVar(variable_name, initial_value)
	glf_initialVars.Add variable_name, initial_value
End Sub

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
			Dim split_key : split_key = Split(line, ":")
            stepTime = Trim(split_key(1))
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
    'msgbox Len(stepLights)
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

		If UBound(showStep.ShowsInStep().Keys()) > -1 Then
			Dim show_item
       	
			Dim show_items : show_items = showStep.ShowsInStep().Items()
        	For Each show_item in show_items
				If Not glf_cached_shows.Exists(show_item.Key & "__" & show_item.InternalCacheId) Then
					Dim cached_show
					cached_show = Glf_ConvertShow(show_item.Show, show_item.Tokens)
					glf_cached_shows.Add show_item.Key & "__" & show_item.InternalCacheId, cached_show
				End If
			Next 
		End If

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
			Dim fadeMs : fadeMs = ""
			Dim fade_duration : fade_duration = -1
			Dim intensity
			Dim step_number : step_number = -1
			Dim localLightsSet : Set localLightsSet = CreateObject("Scripting.Dictionary")
			If IsNull(Glf_IsToken(lightParts(1))) Then
				intensity = lightParts(1)
			Else
				intensity = tokens(Glf_IsToken(lightParts(1)))
			End If
			If Ubound(lightParts) >= 2 Then

				If IsNull(Glf_IsToken(lightParts(2))) Then
					lightColor = lightParts(2)
				Else
					lightColor = tokens(Glf_IsToken(lightParts(2)))
				End If
				If UBound(lightParts) = 3 Then
					If IsNull(Glf_IsToken(lightParts(3))) Then
						fade_duration = lightParts(3)
					Else
						fade_duration = tokens(Glf_IsToken(lightParts(3)))
					End If
					step_number = fade_duration * 0.01
					step_number = Round(step_number, 0) + 2
					If step_number<4 Then
						step_number = 4
					End If
					fadeMs = "|" & step_number
				End If
			End If

			Dim resolved_light_name

			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And Not glf_lightNames.Exists(lightParts(0)) Then
					tagLights = glf_lightTags("T_"&lightParts(0)).Keys()
					resolved_light_name = "T_"&lightParts(0)
					For Each tagLight in tagLights
						If UBound(lightParts) >=1 Then
							seqArray(x) = tagLight & "|"&intensity&"|" & AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						Else
							seqArray(x) = tagLight & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						End If
						If Not localLightsSet.Exists(tagLight) Then
							localLightsSet.Add tagLight, True
						End If
						If Not lightsInShow.Exists(tagLight) Then
							lightsInShow.Add tagLight, True
						End If
						x=x+1
					Next
				Else
					If IsNull(token) Then
						resolved_light_name = lightParts(0)
						If UBound(lightParts) >= 1 Then
							seqArray(x) = lightParts(0) & "|"&intensity&"|"&AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						Else
							seqArray(x) = lightParts(0) & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						End If
						If Not localLightsSet.Exists(lightParts(0)) Then
							localLightsSet.Add lightParts(0), True
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
							resolved_light_name = "T_"&tokens(token)
							For Each tagLight in tagLights
								If UBound(lightParts) >=1 Then
									seqArray(x) = tagLight & "|"&intensity&"|"&AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
								Else
									seqArray(x) = tagLight & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
								End If
								If Not localLightsSet.Exists(tagLight) Then
									localLightsSet.Add tagLight, True
								End If
								If Not lightsInShow.Exists(tagLight) Then
									lightsInShow.Add tagLight, True
								End If
								x=x+1
							Next
						Else
							resolved_light_name = tokens(token)
							If UBound(lightParts) >= 1 Then
								seqArray(x) = tokens(token) & "|"&intensity&"|"&AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
							Else
								seqArray(x) = tokens(token) & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
							End If
							If Not localLightsSet.Exists(tokens(token)) Then
								localLightsSet.Add tokens(token), True
							End If
							If Not lightsInShow.Exists(tokens(token)) Then
								lightsInShow.Add tokens(token), True
							End If
							x=x+1
						End If
					End If
				End If

				'Generate a fake show for the fade in the format light from ? color to lightColor over x steps
				If fadeMs <> "" Then
					Dim fade_seq, i, step_duration,cached_rgb_seq
					 
					Dim cache_name : cache_name = "fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number
					If Not glf_cached_shows.Exists(cache_name & "__-1") Then
						'MsgBox cache_name
						Dim fade_show : Set fade_show = CreateGlfShow(cache_name)
						
					
						step_duration = (fade_duration / step_number)/1000
						For i=1 to step_number

							Dim lightsArr
							ReDim lightsArr(UBound(localLightsSet.Keys))
							Dim localLightItem, k
							k=0
							For Each localLightItem in localLightsSet.Keys()
								cached_rgb_seq = Glf_FadeRGB("FF0000", AdjustHexColor(lightColor, intensity), step_number)
								ReDim fade_seq(step_number - 1)
								Dim j
								For j = 0 To UBound(fade_seq)
									fade_seq(j) = localLightItem & "|100|" & cached_rgb_seq(j)
								Next
								lightsArr(k) = fade_seq(i-1)
								k=k+1
							Next
							With fade_show
								With .AddStep(Null, Null, step_duration)
									.Lights = lightsArr
								End With
							End With
						Next
						cached_show = Glf_ConvertShow(fade_show, Null)
						'msgbox "Converted show: " & cache_name & ", steps: " & ubound(cached_show(0)) & ". Fade Replacements: " & ubound(cached_rgb_seq)
						glf_cached_shows.Add cache_name & "__-1", cached_show
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
	Private m_raw, m_value, m_isGetRef, m_isPlayerState, m_playerStateValue, m_playerStatePlayer, m_isDeviceState, m_deviceStateDeviceType, m_deviceStateDeviceName, m_deviceStateDeviceAttr
  
    Public Property Get Value() 
		If m_isGetRef = True Then
			Value = GetRef(m_value)(Null)
		Else
			Value = m_value
		End If
	End Property

    Public Property Get Raw() : Raw = m_raw : End Property
	Public Property Get RawMValue() : RawMValue = m_value : End Property

	Public Property Get IsPlayerState() : IsPlayerState = m_isPlayerState : End Property
	Public Property Get PlayerStateValue() : PlayerStateValue = m_playerStateValue : End Property		
	Public Property Get PlayerStatePlayer() : PlayerStatePlayer = m_playerStatePlayer : End Property		

	Public Property Get IsDeviceState() : IsDeviceState = m_isDeviceState : End Property
	Public Property Get DeviceStateEvent() : DeviceStateEvent = m_deviceStateDeviceType & "_" & m_deviceStateDeviceName & "_" & m_deviceStateDeviceAttr : End Property
		
	Public default Function init(input)
        m_raw = input
        Dim parsedInput : parsedInput = Glf_ParseInput(input)
		Dim playerState : playerState = Glf_CheckForGetPlayerState(input)
		m_isPlayerState = playerState(0)
		m_playerStateValue = playerState(1)
		m_playerStatePlayer = playerState(2)
		Dim deviceState : deviceState = Glf_CheckForDeviceState(input)
		m_isDeviceState = deviceState(0)
		m_deviceStateDeviceType = deviceState(1)
		m_deviceStateDeviceName = deviceState(2)
		m_deviceStateDeviceAttr = deviceState(3)
		
        m_value = parsedInput(0)
        m_isGetRef = parsedInput(2)
	    Set Init = Me
	End Function

End Class

Function Glf_FadeRGB(color1, color2, steps)
	If glf_cached_rgb_fades.Exists(color1&"_"&color2&"_"&CStr(steps)) Then
		Glf_FadeRGB = glf_cached_rgb_fades(color1&"_"&color2&"_"&CStr(steps))
		Exit Function
	End If

	Dim cached_rgb_seq : cached_rgb_seq = Array()
	Dim r1, g1, b1, r2, g2, b2, c1,c2
	Dim i
	Dim r, g, b
	c1 = clng( RGB( Glf_HexToInt(Left(color1, 2)), Glf_HexToInt(Mid(color1, 3, 2)), Glf_HexToInt(Right(color1, 2)))  )
	c2 = clng( RGB( Glf_HexToInt(Left(color2, 2)), Glf_HexToInt(Mid(color2, 3, 2)), Glf_HexToInt(Right(color2, 2)))  )
	
	r1 = c1 Mod 256
	g1 = (c1 \ 256) Mod 256
	b1 = (c1 \ (256 * 256)) Mod 256

	r2 = c2 Mod 256
	g2 = (c2 \ 256) Mod 256
	b2 = (c2 \ (256 * 256)) Mod 256

	ReDim cached_rgb_seq(steps - 1)
	Dim rgb_color
	For i = 0 To steps - 1
		r = r1 + (r2 - r1) * i / (steps - 1)
		g = g1 + (g2 - g1) * i / (steps - 1)
		b = b1 + (b2 - b1) * i / (steps - 1)
		rgb_color = Glf_RGBToHex(CInt(r), CInt(g), CInt(b))
		cached_rgb_seq(i) = rgb_color
	Next
	glf_cached_rgb_fades.Add color1&"_"&color2&"_"&CStr(steps), cached_rgb_seq	
	Glf_FadeRGB = cached_rgb_seq
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

With CreateGlfShow("fade_led_color")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|(color)|(fade)")
	End With
End With

With CreateGlfShow("fade_rgb_test")
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|ff0000|(fade)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|00ff00|(fade)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|0000ff|(fade)")
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
	If hexColor = "stop" Then
		AdjustHexColor = "stop"
		Exit Function
	End If
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

