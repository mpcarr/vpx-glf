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
Dim bcpExeName : bcpExeName = ""
Dim glf_monitor_player_vars : glf_monitor_player_vars = false
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0
Dim glf_SeqCount : glf_SeqCount = 0
Dim glf_max_dispatch : glf_max_dispatch = 25
Dim glf_max_lightmap_sync : glf_max_lightmap_sync = -1
Dim glf_max_lightmap_sync_enabled : glf_max_lightmap_sync_enabled = False
Dim glf_max_lights_test : glf_max_lights_test = 0

Dim glf_master_volume : glf_master_volume = 0.8


Dim glf_troughSize : glf_troughSize = tnob
Dim glf_lastTroughSw : glf_lastTroughSw = Null
Dim glf_game
With GlfGameSettings()
	.BallsPerGame = 3
End With
Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False
Dim glf_debug_level : glf_debug_level = "Info"


Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController(args)
    Set bcpController = (new GlfVpxBcpController)(bcpPort, bcpExeName)
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

Public Sub Glf_Init()
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
			showsYaml = "#show_version=6" & vbCrLf & vbCrLf
			showsYaml = showsYaml + device.ToYaml()
			Set TxtFileStream = fso.OpenTextFile(showsFolder & "\" & device.Name & ".yaml", 2, True)
			TxtFileStream.WriteLine showsYaml
			TxtFileStream.Close
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
	glf_debugLog.WriteToLog "Code String", glf_codestr

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
			DispatchRelayPinEvent "request_to_start_game", True
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
'     pattern = "devices\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
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
    startPos = InStr(outputString, "devices.")

    Do While startPos > 0
        ' Find the end of the match by looking for three dots after "devices."
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
            startPos = InStr(startPos + Len(replacement), outputString, "devices.")
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

    pos = InStr(inputString, "devices.")
    If pos > 0 Then
        segStart = pos + Len("devices.")
        
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


Function EnableGlfBallSearch()
	Dim ball_search : Set ball_search = (new GlfBallSearch)()
    With CreateGlfMode("glf_ball_search", 100)
        .StartEvents = Array("reset_complete")

        With .TimedSwitches("flipper_cradle")
            .Switches = Array("s_left_flipper", "s_right_flipper")
            .Time = 3000
            .EventsWhenActive = Array("flipper_cradle")
            .EventsWhenReleased = Array("flipper_release")
        End With
    End With
    glf_ballsearch_enabled = True
	Set EnableGlfBallSearch = ball_search
End Function

Class GlfBallSearch

    Private m_debug
    Private m_timeout
    Private m_search_interval
    Private m_ball_search_wait_after_iteration
    Private m_phase
    Private m_devices
    Private m_current_device_type

    Public Property Get GetValue(value)
        'Select Case value
            'Case   
        'End Select
        GetValue = True
    End Property

    Public Property Let Timeout(value): Set m_timeout = CreateGlfInput(value): End Property
    Public Property Get Timeout(): Timeout = m_timeout.Value(): End Property

    Public Property Let SearchInterval(value): Set m_search_interval = CreateGlfInput(value): End Property
    Public Property Get SearchInterval(): SearchInterval = m_search_interval.Value(): End Property

    Public Property Let BallSearchWaitAfterIteration(value): Set m_ball_search_wait_after_iteration = CreateGlfInput(value): End Property
    Public Property Get BallSearchWaitAfterIteration(): BallSearchWaitAfterIteration = m_ball_search_wait_after_iteration.Value(): End Property

    Public Property Let Debug(value)
        m_debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init()
        Set m_timeout = CreateGlfInput(15000)
        Set m_search_interval = CreateGlfInput(150)
        Set m_ball_search_wait_after_iteration = CreateGlfInput(10000)
        m_phase = 0
        m_devices = Array()
        m_current_device_type = Empty
        Set glf_ballsearch = Me
        SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
        AddPinEventListener "flipper_cradle", "ball_search_flipper_cradle", "BallSearchHandler", 30, Array("stop", Me)
        AddPinEventListener "flipper_release", "ball_search_flipper_cradle", "BallSearchHandler", 30, Array("reset", Me)
        Set Init = Me
    End Function

    Public Sub Start(phase)
        Dim ball_hold, held_balls
        held_balls = 0
        For Each ball_hold in glf_ball_holds.Items()
            held_balls = held_balls + ball_hold.GetValue("balls_held")
        Next
        If glf_gameStarted = True And glf_BIP > 0 And (glf_BIP-held_balls)>0 And glf_plunger.HasBall() = False Then
            m_phase = phase
            glf_last_switch_hit_time = 0
            'Fire all auto fire devices, slings, pops.
            m_devices = glf_autofiredevices.Items()
            m_current_device_type = "autofire"
            DispatchPinEvent "ball_search_started", Null
            If UBound(m_devices) > -1 Then
                m_devices(0).BallSearch(m_phase)
                SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
            End If
        Else
            SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
        End If
    End Sub

    Public Sub NextDevice(device_index)
        If UBound(m_devices) > device_index Then
            m_devices(device_index+1).BallSearch(m_phase)
            SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, device_index+1), Null), m_search_interval.Value
        Else
            If m_current_device_type = "autofire" Then
                m_devices = glf_ball_devices.Items()
                m_current_device_type = "balldevices"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            ElseIf m_current_device_type = "balldevices" Then
                m_devices = glf_drop_targets.Items()
                m_current_device_type = "droptargets"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            ElseIf m_current_device_type = "droptargets" Then
                m_devices = glf_diverters.Items()
                m_current_device_type = "diverters"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            Else
                m_current_device_type = Empty
                If m_phase < 3 Then
                    Start m_phase+1
                Else
                    m_phase = 0
                    SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
                End If
            End If
        End If
    End Sub

    Public Sub Reset()
        RemoveDelay "ball_search_next_device"
        DispatchPinEvent "ball_search_stopped", Null
        m_phase = 0
        SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
    End Sub

    Public Sub StopBallSearch()
        DispatchPinEvent "ball_search_stopped", Null
        RemoveDelay "ball_search_next_device"
        m_phase = 0
        RemoveDelay "ball_search"
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function BallSearchHandler(args)
    Dim ownProps, kwargs
    ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ball_search : Set ball_search = ownProps(1)
    Select Case evt
        Case "start"
            ball_search.Start 1
        Case "reset":
            ball_search.Reset
        Case "stop":
            ball_search.StopBallSearch
        Case "next_device"
            ball_search.NextDevice ownProps(2)
    End Select
End Function



'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************

Class GlfVpxBcpController

    Private m_bcpController, m_connected, m_mode_list

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_controller.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        useBcp = True
        m_mode_list = ""
        AddPinEventListener "player_added", "bcp_player_added", "GlfVpxBcpControllerEventHandler", 50, Array("player_added")
        AddPinEventListener "next_player", "bcp_player_next_player", "GlfVpxBcpControllerEventHandler", 50, Array("next_player")
        AddPinEventListener "game_started", "bcp_player_next_player", "GlfVpxBcpControllerEventHandler", 50, Array("next_player")
        AddPinEventListener "ball_ended", "bcp_player_ball_end", "GlfVpxBcpControllerEventHandler", 50, Array("ball_end")
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
    
    Public Sub PlaySlide(slide, context, calling_context, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""" & slide & """: {""action"": ""play"", ""expire"": 0}}, ""context"": """ & context & """, ""calling_context"": """ & calling_context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub ModeList()
        If m_connected Then
            If m_mode_list <> glf_running_modes Then
                m_bcpController.Send "mode_list?json={""running_modes"": ["&glf_running_modes&"]}"
                m_mode_list = glf_running_modes
            End If
        End If
    End Sub

    Public Sub SlidesClear(context)
        If m_connected Then
            m_bcpController.Send "trigger?name=slides_clear&context=" & context
        End If
    End Sub

    Public Sub ModeStart(name, priority)
        If m_connected Then
            m_bcpController.Send "mode_start?priority=int:" & priority & "&name=" & name
        End If
    End Sub

    Public Sub ModeStop(name)
        If m_connected Then
            m_bcpController.Send "mode_stop?name=" & name
        End If
    End Sub


    Public Sub SendPlayerVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & Getglf_currentPlayerNumber+1
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
                retValue = value
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
            useBcp = False
        End If
    End Sub
End Class

Sub Glf_BcpSendPlayerVar(args)
    If useBcp=False Then
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
    If useBcp=False Then
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
                        glf_monitor_player_vars = True
                        'AddPlayerStateEventListener "score", "bcp_player_var_score_0", 0, "Glf_BcpSendPlayerVar", 1000, Null
                        'AddPlayerStateEventListener "current_ball", "bcp_player_var_ball_0", 0, "Glf_BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    eventName = message.GetValue("event")
            End Select
        Next
    End If
    bcpController.ModeList()
End Sub

Function GlfVpxBcpControllerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Select Case evt
        Case "player_added"
            bcpController.Send "player_added?player_num=int:" & kwargs("num")
        Case "next_player"
            Dim p_num : p_num = Getglf_currentPlayerNumber() + 1
            bcpController.Send "player_turn_start?player_num=int:" & p_num
            bcpController.Send "ball_start?player_num=int:" & p_num & "&ball=int:" & GetPlayerState("ball")
            bcpController.SendPlayerVariable "number", p_num, p_num-1
            'bcpController.SendPlayerVariable "number", 1, 0
        Case "ball_end"
            bcpController.Send "ball_end"
    End Select
    If IsObject(args(1)) Then
        Set GlfVpxBcpControllerEventHandler = kwargs
    Else
        GlfVpxBcpControllerEventHandler = kwargs
    End If
    
End Function

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************


'*****************************************************************************************************************************************
'  Glf Monitor Bcp Controller
'*****************************************************************************************************************************************

Class GlfMonitorBcpController

    Private m_bcpController, m_connected, m_isInMonitor

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_controller.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        If Err Then MsgBox("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

    Public Function IsInMonitior()
        If m_connected = True And m_isInMonitor = True Then
            IsInMonitior=True
        Else
            IsInMonitior=False
        End If
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
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""monitor"": {""action"": ""play"", ""expire"": 0}}, ""context"": """", ""priority"": 1}"
            
            Dim mode
            For Each mode in glf_modes.Items()
                Glf_MonitorModeUpdate mode
            Next
            m_isInMonitor = True
        End If
	End Sub

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
        End If
    End Sub
End Class

Sub Glf_MonitorModeUpdate(mode)
    If glf_hasDebugController = False Then
        Exit Sub
    End If
    glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """&mode.Status&""", ""debug"": " & mode.IsDebug & "},"
    Dim config_item
    For Each config_item in mode.BallSavesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.CountersItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.TimersItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.MultiballLocksItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.MultiballsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ModeShots()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ShotGroupsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.BallHoldsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.SequenceShotsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ExtraBallsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ComboSwitchesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.TimedSwitchesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ModeStateMachines()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    If Not IsNull(mode.LightPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.LightPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.LightPlayer.Name & """},"
    End If
    If Not IsNull(mode.EventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.EventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.EventPlayer.Name & """},"
    End If
    If Not IsNull(mode.TiltConfig) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.TiltConfig.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.TiltConfig.Name & """},"
    End If
    If Not IsNull(mode.QueueEventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.QueueEventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.QueueEventPlayer.Name & """},"
    End If
    If Not IsNull(mode.QueueRelayPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.QueueRelayPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.QueueRelayPlayer.Name & """},"
    End If
    If Not IsNull(mode.RandomEventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.RandomEventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.RandomEventPlayer.Name & """},"
    End If
    If Not IsNull(mode.ShowPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.ShowPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.ShowPlayer.Name & """},"
    End If
    If Not IsNull(mode.SegmentDisplayPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.SegmentDisplayPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.SegmentDisplayPlayer.Name & """},"
    End If
    If Not IsNull(mode.VariablePlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.VariablePlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.VariablePlayer.Name & """},"
    End If
    If Not IsNull(mode.DOFPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.DOFPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.DOFPlayer.Name & """},"
    End If
    If Not IsNull(mode.SlidePlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.SlidePlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.SlidePlayer.Name & """},"
    End If
End Sub

Sub Glf_MonitorPlayerStateUpdate(key, value)
    If glf_hasDebugController = False Then
        Exit Sub
    End If    
    glf_monitor_player_state = glf_monitor_player_state & "{""key"": """&key&""", ""value"": """&value&"""},"
End Sub

Sub Glf_MonitorEventStream(label, message)
    If glf_hasDebugController = False Then
        Exit Sub
    End If
    glf_monitor_event_stream = glf_monitor_event_stream & "{""label"": """&label&""", ""message"": """&message&"""},"
End Sub


Sub Glf_MonitorBcpUpdate()
    If glf_hasDebugController = False Then
        Exit Sub
    End If

    'Send Updates
    If glf_debugBcpController.IsInMonitior Then
        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_player_state"",""changes"": ["&glf_monitor_player_state&"]}"
        glf_monitor_player_state = ""

        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_monitor_modes"",""changes"": ["&glf_monitor_modes&"]}"
        glf_monitor_modes = ""

        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_event_stream"",""changes"": ["&glf_monitor_event_stream&"]}"
        glf_monitor_event_stream = ""
    End If
    

    Dim messages : messages = glf_debugBcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter, eventName
        For Each message in messages
            debug.print(message.Command)
            Select Case message.Command
                case "hello"
                    glf_debugBcpController.Reset
                case "trigger"
                    eventName = message.GetValue("name")
                    Dim mode_name, device_name
                    debug.print eventName
                    If eventName = "glf_monitor_debug_mode" Then
                        mode_name = message.GetValue("mode")
                        debug.print mode_name
                        If Not IsNull(GlfModes(mode_name)) Then
                            debug.print("got mode")
                            If GlfModes(mode_name).IsDebug = 1 Then
                                debug.print("Turning off debug")
                                GlfModes(mode_name).Debug = False
                            Else
                                debug.print("Turning on debug")
                                GlfModes(mode_name).Debug = True
                            End If
                        End If
                    End If
                    If eventName = "glf_monitor_debug_mode_device" Then
                        mode_name = message.GetValue("mode")
                        device_name = message.GetValue("mode_device")
                        device_name = Replace(device_name, mode_name & "_", "")
                        debug.print mode_name
                        debug.print device_name
                        If Not IsNull(GlfModes(mode_name)) Then
                            debug.print("got mode")
                            Dim config_item, mode, is_debug
                            is_debug = message.GetValue("debug")
                            debug.print is_debug
                            If is_debug = "bool:true" Then
                                is_debug = True
                            Else
                                is_debug = False
                            End If
                            Set mode = GlfModes(mode_name)
                            For Each config_item in mode.BallSavesItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.CountersItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.TimersItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.MultiballLocksItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.MultiballsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ModeShots()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ShotGroupsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.BallHoldsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.SequenceShotsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ExtraBallsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ComboSwitchesItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.TimedSwitchesItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ModeStateMachines()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            If Not IsNull(mode.LightPlayer) Then
                                If mode.LightPlayer.Name = device_name Then : mode.LightPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.EventPlayer) Then
                                If mode.EventPlayer.Name = device_name Then : mode.EventPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.TiltConfig) Then
                                If mode.TiltConfig.Name = device_name Then : mode.TiltConfig.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.RandomEventPlayer) Then
                                If mode.RandomEventPlayer.Name = device_name Then : mode.RandomEventPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.ShowPlayer) Then
                                If mode.ShowPlayer.Name = device_name Then : mode.ShowPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.SegmentDisplayPlayer) Then
                                If mode.SegmentDisplayPlayer.Name = device_name Then : mode.SegmentDisplayPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.VariablePlayer) Then
                                If mode.VariablePlayer.Name = device_name Then : mode.VariablePlayer.Debug = is_debug : End If
                            End If
                        End If
                    End If
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************a
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



Class GlfBallHold

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_enabled
    Private m_balls_to_hold
    Private m_hold_devices
    Private m_balls_held
    Private m_hold_queue
    Private m_release_all_events
    Private m_release_one_events
    Private m_release_one_if_full_events

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "balls_held":
                GetValue = m_balls_held
        End Select
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

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
            m_release_all_events.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Get ReleaseOneEvents(): Set ReleaseOneEvents = m_release_one_events: End Property
    Public Property Let ReleaseOneEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_one_events.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Get ReleaseOneIfFullEvents(): Set ReleaseOneIfFullEvents = m_release_one_if_full_events: End Property
    Public Property Let ReleaseOneIfFullEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_one_if_full_events.Add newEvent.Raw, newEvent
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
        ReleaseAllEvents = Array("tilt")
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
            AddPinEventListener m_release_all_events(evt).EventName, m_mode & "_" & name & "_release_all", "BallHoldsEventHandler", m_priority+m_release_all_events(evt).Priority, Array("release_all", me, m_release_all_events(evt))
        Next
        For Each evt in m_release_one_events.Keys
            AddPinEventListener m_release_one_events(evt).EventName, m_mode & "_" & name & "_release_one", "BallHoldsEventHandler", m_priority+m_release_one_events(evt).Priority, Array("release_one", me, m_release_one_events(evt))
        Next
        For Each evt in m_release_one_if_full_events.Keys
            AddPinEventListener m_release_one_if_full_events(evt).EventName, m_mode & "_" & name & "_release_one_if_full", "BallHoldsEventHandler", m_priority+m_release_one_if_full_events(evt).Priority, Array("release_one_if_full", me, m_release_one_if_full_events(evt))
        Next
    End Sub

    Public Sub Disable()
        m_enabled = False
        ReleaseAll()
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
            Log "Cannot hold balls. Hold is full."
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
        Log "Held " & balls_to_hold & " balls"

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

        Log "Releasing up to " & balls_to_release & " balls from hold"
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

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt, x
        yaml = "  " & Replace(m_name, "ball_hold_", "") & ":" & vbCrLf
        Dim enable_events_keys : enable_events_keys = m_base_device.EnableEvents().Keys
        Dim enable_events : Set enable_events = m_base_device.EnableEvents()
        If UBound(enable_events_keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in enable_events_keys
                yaml = yaml & enable_events(key).Raw
                If x <> UBound(enable_events_keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        Dim disable_events_keys : disable_events_keys = m_base_device.DisableEvents().Keys
        Dim disable_events : Set disable_events = m_base_device.DisableEvents()
        If UBound(disable_events_keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in disable_events_keys
                yaml = yaml & disable_events(key).Raw
                If x <> UBound(disable_events_keys) Then
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
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            ball_hold.ReleaseAll
        Case "release_one"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            ball_hold.ReleaseBalls 1
        Case "release_one_if_full"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
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
    Private m_priority
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
    private m_base_device
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AutoLaunch(): AutoLaunch = m_auto_launch: End Property
    Public Property Let ActiveTime(value) : m_active_time = Glf_ParseInput(value) : End Property
    Public Property Let GracePeriod(value) : m_grace_period = Glf_ParseInput(value) : End Property
    Public Property Let HurryUpTime(value) : m_hurry_up_time = Glf_ParseInput(value) : End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property

    Public Property Let TimerStartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_timer_start_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let AutoLaunch(value) : m_auto_launch = value : End Property
    Public Property Let BallsToSave(value) : m_balls_to_save = value : End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property

	Public default Function init(name, mode)
        m_name = "ball_save_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
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
        Set m_base_device = (new GlfBaseModeDevice)(mode, "ball_save", Me)
	    Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_timer_start_events.Keys
            AddPinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start", "BallSaveEventHandler", m_priority+m_timer_start_events(evt).Priority, Array("timer_start", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_timer_start_events.Keys
            RemovePinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start"
        Next
    End Sub

    Public Sub Enable()
        If m_enabled = True Then
            Exit Sub
        End If
        m_enabled = True
        m_saving_balls = m_balls_to_save
        Log "Enabling. Auto launch: "&m_auto_launch&", Balls to save: "&m_balls_to_save
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "BallSaveEventHandler", 1000, Array("drain", Me)
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
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
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
            Dim active_time : active_time = GetRef(m_active_time(0))(Null)
            Dim grace_period, hurry_up_time
            If Not IsNull(m_grace_period) Then
                grace_period = GetRef(m_grace_period(0))(Null)
            Else
                grace_period = 0
            End If
            If Not IsNull(m_hurry_up_time) Then
                hurry_up_time = GetRef(m_hurry_up_time(0))(Null)
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
        
        ' Add disable_events if any exist
        If m_disable_events.Count > 0 Then
            yaml = yaml & "    disable_events: "
            x = 0
            For Each evt in m_disable_events.Keys
                yaml = yaml & m_disable_events(evt).Raw
                If x <> UBound(m_disable_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        
        ' Add timer_stop_events if any exist
        If m_timer_stop_events.Count > 0 Then
            yaml = yaml & "    timer_stop_events: "
            x = 0
            For Each evt in m_timer_stop_events.Keys
                yaml = yaml & m_timer_stop_events(evt).Raw
                If x <> UBound(m_timer_stop_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        
        ' Add timer_complete_events if any exist
        If m_timer_complete_events.Count > 0 Then
            yaml = yaml & "    timer_complete_events: "
            x = 0
            For Each evt in m_timer_complete_events.Keys
                yaml = yaml & m_timer_complete_events(evt).Raw
                If x <> UBound(m_timer_complete_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        
        ' Add ball_save_events if any exist
        If m_ball_save_events.Count > 0 Then
            yaml = yaml & "    ball_save_events: "
            x = 0
            For Each evt in m_ball_save_events.Keys
                yaml = yaml & m_ball_save_events(evt).Raw
                If x <> UBound(m_ball_save_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        
        ' Add ball_save_complete_events if any exist
        If m_ball_save_complete_events.Count > 0 Then
            yaml = yaml & "    ball_save_complete_events: "
            x = 0
            For Each evt in m_ball_save_complete_events.Keys
                yaml = yaml & m_ball_save_complete_events(evt).Raw
                If x <> UBound(m_ball_save_complete_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        
        ' Add debug setting if enabled
        If m_debug Then
            yaml = yaml & "    debug: true" & vbCrLf
        End If
        
        ToYaml = yaml
    End Function
End Class

Function BallSaveEventHandler(args)
    Dim ownProps, ballsToSave : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    ballsToSave = args(1) 
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
            If glf_plunger.HasBall = False And ballInReleasePostion = True  And glf_plunger.IncomingBalls = 0  Then
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


Class GlfComboSwitches

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_switch_1
    Private m_switch_2
    Private m_events_when_both
    Private m_events_when_inactive
    Private m_events_when_one
    Private m_events_when_switch_1
    Private m_events_when_switch_2
    Private m_hold_time
    Private m_max_offset_time
    Private m_release_time

    Private m_switch_1_active
    Private m_switch_2_active

    Private m_switch_state

    Public Property Get Name() : Name = m_name : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get GetValue(value)
        Select Case value
            'Case ""
            '    GetValue = m_ticks
        End Select
    End Property

    Public Property Let Switch1(value): m_switch_1 = value: End Property
    Public Property Let Switch2(value): m_switch_2 = value: End Property
    Public Property Let HoldTime(value): Set m_hold_time = CreateGlfInput(value): End Property
    Public Property Let MaxOffsetTime(value): Set m_max_offset_time = CreateGlfInput(value): End Property
    Public Property Let ReleaseTime(value): Set m_release_time = CreateGlfInput(value): End Property
    Public Property Let EventsWhenBoth(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_both.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenInactive(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_inactive.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenOne(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_one.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenSwitch1(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_switch_1.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenSwitch2(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_switch_2.Add newEvent.Raw, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = "combo_switch_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
    
        m_switch_1 = Empty
        m_switch_2 = Empty
        Set m_events_when_both = CreateObject("Scripting.Dictionary")
        Set m_events_when_inactive = CreateObject("Scripting.Dictionary")
        Set m_events_when_one = CreateObject("Scripting.Dictionary")
        Set m_events_when_switch_1 = CreateObject("Scripting.Dictionary")
        Set m_events_when_switch_2 = CreateObject("Scripting.Dictionary")
        Set m_hold_time = CreateGlfInput(0)
        Set m_max_offset_time = CreateGlfInput(-1)
        Set m_release_time = CreateGlfInput(0)

        m_switch_1_active = 0
        m_switch_2_active = 0

        m_switch_state = Empty
        Set m_base_device = (new GlfBaseModeDevice)(mode, "combo_switch", Me)

        glf_combo_switches.Add name, Me

        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating Combo Switch"
        AddPinEventListener m_switch_1 & "_active" , m_name & "_switch1_active" , "ComboSwitchEventHandler", m_priority, Array("switch1_active", Me, m_switch_1)
        AddPinEventListener m_switch_1 & "_inactive" , m_name & "_switch1_inactive" , "ComboSwitchEventHandler", m_priority, Array("switch1_inactive", Me, m_switch_1)
        AddPinEventListener m_switch_2 & "_active" , m_name & "_switch2_active" , "ComboSwitchEventHandler", m_priority, Array("switch2_active", Me, m_switch_2)
        AddPinEventListener m_switch_2 & "_inactive" , m_name & "_switch2_inactive" , "ComboSwitchEventHandler", m_priority, Array("switch2_inactive", Me, m_switch_2)
    End Sub

    Public Sub Deactivate()
        Log "Deactivating Combo Switch"
        RemovePinEventListener m_switch_1 & "_active" , m_name & "_switch1_active"
        RemovePinEventListener m_switch_1 & "_inactive" , m_name & "_switch1_inactive"
        RemovePinEventListener m_switch_2 & "_active" , m_name & "_switch2_active"
        RemovePinEventListener m_switch_2 & "_inactive" , m_name & "_switch2_inactive"
        RemoveDelay m_name & "_" & "switch_1_inactive"
        RemoveDelay m_name & "_" & "switch_2_inactive"
        RemoveDelay m_name & "_" & "switch_1_active"
        RemoveDelay m_name & "_" & "switch_1_active"
        RemoveDelay m_name & "_" & "switch_2_only"
        RemoveDelay m_name & "_" & "switch_1_only"
    End Sub

    Public Sub Switch1WentActive(switch_name)
        Log "switch_1 just went active"
        RemoveDelay m_name & "_" & "switch_1_inactive"

        If m_switch_1_active > 0 Then
            Exit Sub
        End If

        If m_hold_time.Value() = 0 Then
            ActivateSwitches1 switch_name
        Else
            SetDelay m_name & "_switch_1_active", "ComboSwitchEventHandler", Array(Array("activate_switch1", Me, switch_name),Null), m_hold_time.Value()
        End If
    End Sub

    Public Sub Switch2WentActive(switch_name)
        Log "switch_2 just went active"
        RemoveDelay m_name & "_" & "switch_2_inactive"

        If m_switch_2_active > 0 Then
            Exit Sub
        End If

        If m_hold_time.Value() = 0 Then
            ActivateSwitches2 switch_name
        Else
            SetDelay m_name & "_switch_2_active", "ComboSwitchEventHandler", Array(Array("activate_switch2", Me, switch_name),Null), m_hold_time.Value()
        End If
    End Sub

    Public Sub Switch1WentInactive(switch_name)
        Log "switch_1 just went inactive"
        RemoveDelay m_name & "_" & "switch_1_active"

        If m_release_time.Value() = 0 Then
            ReleaseSwitch1 switch_name
        Else
            SetDelay m_name & "_switch_1_inactive", "ComboSwitchEventHandler", Array(Array("release_switch1", Me, switch_name),Null), m_release_time.Value()
        End If
    End Sub

    Public Sub Switch2WentInactive(switch_name)
        Log "switch_2 just went inactive"
        RemoveDelay m_name & "_" & "switch_2_active"

        If m_release_time.Value() = 0 Then
            ReleaseSwitch2 switch_name
        Else
            SetDelay m_name & "_switch_2_inactive", "ComboSwitchEventHandler", Array(Array("release_switch2", Me, switch_name),Null), m_release_time.Value()
        End If
    End Sub

    Public Sub ActivateSwitches1(switch_name)
        Log "Switch_1 has passed the hold time and is now active"
        m_switch_1_active = gametime
        RemoveDelay m_name & "_" & "switch_2_only"

        If m_switch_2_active > 0 Then
            If m_max_offset_time.Value() >= 0 And (m_switch_1_active - m_switch_2_active) > m_max_offset_time.Value() Then
                Log "Switches_2 is active, but the max_offset_time=" & m_max_offset_time.Value() & " which is larger than when a Switches_1 switch was first activated, so the state will not switch to both"
                Exit Sub
            End If
            PostSwitchStateEvents "both"
        ElseIf m_max_offset_time.Value()>=0 Then
            SetDelay m_name & "_switch_1_only", "ComboSwitchEventHandler", Array(Array("switch_1_only", Me, switch_name),Null), max_offset_time.Value()
        End If
    End Sub

    Public Sub ActivateSwitches2(switch_name)
        Log "Switch_2 has passed the hold time and is now active"
        m_switch_2_active = gametime
        RemoveDelay m_name & "_" & "switch_1_only"

        If m_switch_1_active > 0 Then
            If m_max_offset_time.Value() >= 0 And (m_switch_2_active - m_switch_1_active) > m_max_offset_time.Value() Then
                Log "Switches_1 is active, but the max_offset_time=" & m_max_offset_time.Value() & " which is larger than when a Switches_2 switch was first activated, so the state will not switch to both"
                Exit Sub
            End If
            PostSwitchStateEvents "both"
        ElseIf m_max_offset_time.Value()>=0 Then
            SetDelay m_name & "_switch_2_only", "ComboSwitchEventHandler", Array(Array("switch_2_only", Me, switch_name),Null), max_offset_time.Value()
        End If
    End Sub

    Public Sub ReleaseSwitch1(switch_name)
        Log "Switches_1 has passed the release time and is now released"
        m_switch_1_active = 0

        If m_switch_2_active > 0 And m_switch_state = "both" Then
            PostSwitchStateEvents "one"
        ElseIf m_switch_state = "one" Then
            PostSwitchStateEvents "inactive"
        End If
    End Sub

    Public Sub ReleaseSwitch2(switch_name)
        Log "Switches_2 has passed the release time and is now released"
        m_switch_2_active = 0

        If m_switch_1_active > 0 And m_switch_state = "both" Then
            PostSwitchStateEvents "one"
        ElseIf m_switch_state = "one" Then
            PostSwitchStateEvents "inactive"
        End If
    End Sub

    Public Sub PostSwitchStateEvents(state)
        If m_switch_state = state Then
            Exit Sub
        End If
        m_switch_state = state
        Log "New State " & state

        Dim evt
        Select Case state
            Case "both"
                For Each evt in m_events_when_both.Keys
                    If m_events_when_both(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_both(evt).EventName, Null
                    End If
                Next
            Case "one"
                For Each evt in m_events_when_one.Keys
                    If m_events_when_one(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_one(evt).EventName, Null
                    End If
                Next
            Case "inactive"
                For Each evt in m_events_when_inactive.Keys
                    If m_events_when_inactive(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_inactive(evt).EventName, Null
                    End If
                Next
            Case "switches_1"
                For Each evt in m_events_when_switch_1.Keys
                    If m_events_when_switch_1(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_switch_1(evt).EventName, Null
                    End If
                Next
            Case "switches_2"
                For Each evt in m_events_when_switch_2.Keys
                    If m_events_when_switch_2(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_switch_2(evt).EventName, Null
                    End If
                Next
        End Select

    End Sub

    Public Function ToYaml()
        Dim yaml, key, x
        yaml = "  " & Replace(m_name, "combo_switch_", "") & ":" & vbCrLf
        If Not IsEmpty(m_switch_1) Then
            yaml = yaml & "    switches_1: " & m_switch_1 & vbCrLf
        End If
        If Not IsEmpty(m_switch_2) Then
            yaml = yaml & "    switches_2: " & m_switch_2 & vbCrLf
        End If
        If UBound(m_events_when_both.Keys()) > -1 Then
            yaml = yaml & "    events_when_both: " & vbCrLf
            For Each key in m_events_when_both.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_inactive.Keys()) > -1 Then
            yaml = yaml & "    events_when_inactive: " & vbCrLf
            For Each key in m_events_when_inactive.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_switch_1.Keys()) > -1 Then
            yaml = yaml & "    events_when_switches_1: " & vbCrLf
            For Each key in m_events_when_switch_1.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_switch_2.Keys()) > -1 Then
            yaml = yaml & "    events_when_switches_2: " & vbCrLf
            For Each key in m_events_when_switch_2.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_switch_2.Keys()) > -1 Then
            yaml = yaml & "    events_when_switches_2: " & vbCrLf
            For Each key in m_events_when_switch_2.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If m_hold_time.Raw <> 0 Then
            yaml = yaml & "    hold_time: " & m_hold_time.Raw & vbCrLf
        End If
        If m_max_offset_time.Raw <> 0 Then
            yaml = yaml & "    max_offset_time: " & m_max_offset_time.Raw & vbCrLf
        End If
        If m_release_time.Raw <> 0 Then
            yaml = yaml & "    release_time: " & m_release_time.Raw & vbCrLf
        End If

        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function ComboSwitchEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim combo_switch : Set combo_switch = ownProps(1)
    Dim switch_name : switch_name = ownProps(2)
    Select Case evt
        Case "switch1_active"
            combo_switch.Switch1WentActive switch_name
        Case "switch2_active"
            combo_switch.Switch2WentActive switch_name
        Case "switch1_inactive"
            combo_switch.Switch1WentInactive switch_name
        Case "switch2_inactive"
            combo_switch.Switch2WentInactive switch_name
        Case "activate_switch1"
            combo_switch.ActivateSwitches1 switch_name
        Case "activate_switch2"
            combo_switch.ActivateSwitches2 switch_name
        Case "release_switch1"
            combo_switch.ReleaseSwitch1 switch_name
        Case "release_switch2"
            combo_switch.ReleaseSwitch2 switch_name
        Case "switch_1_only"
            combo_switch.PostSwitchStateEvents "switches_1"
        Case "switch_2_only"
            combo_switch.PostSwitchStateEvents "switches_2"
    End Select
    If IsObject(args(1)) Then
        Set ComboSwitchEventHandler = kwargs
    Else
        ComboSwitchEventHandler = kwargs
    End If
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
            m_enable_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let CountEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_count_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let CountCompleteValue(value) : m_count_complete_value = value : End Property
    Public Property Let DisableOnComplete(value) : m_disable_on_complete = value : End Property
    Public Property Let ResetOnComplete(value) : m_reset_on_complete = value : End Property
    Public Property Let EventsWhenComplete(value) : m_events_when_complete = value : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

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
            If GetPlayerState(m_name & "_state")=False Then
                SetValue 0
            Else
                SetValue GetPlayerState(m_name & "_state")
            End If
        Else
            SetValue 0
        End If
        Dim evt
        For Each evt in m_enable_events.Keys
            AddPinEventListener m_enable_events(evt).EventName, m_name & "_enable", "CounterEventHandler", m_priority+m_enable_events(evt).Priority, Array("enable", Me, evt)
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
            AddPinEventListener m_count_events(evt).EventName, m_name & "_count", "CounterEventHandler", m_priority+m_count_events(evt).Priority, Array("count", Me)
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


Class GlfDofPlayer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "dof_player" : End Property


    Public Property Get EventDOF() : EventDOF = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_dof : Set new_dof = (new GlfDofPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_dof
        Set EventName = new_dof
        
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "dof_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "dof_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_dof_player_play", "DofPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_dof_player_play"
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            Log "Firing DOF Event: " & m_eventValues(evt).DOFEvent & " State: " & m_eventValues(evt).Action
            DOF m_eventValues(evt).DOFEvent, m_eventValues(evt).Action  
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_dof_player", message
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

Function DofPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim dof_player : Set dof_player = ownProps(1)
    Select Case evt
        Case "activate"
            dof_player.Activate
        Case "deactivate"
            dof_player.Deactivate
        Case "play"
            dof_player.Play(ownProps(2))
    End Select
    If IsObject(args(1)) Then
        Set DofPlayerEventHandler = kwargs
    Else
        DofPlayerEventHandler = kwargs
    End If
End Function

Class GlfDofPlayerItem
	Private m_dof_event, m_action
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        Select Case input
            Case "DOF_OFF"
                m_action = 0
            Case "DOF_ON"
                m_action = 1
            Case "DOF_PULSE"
                m_action = 2
        End Select
    End Property

    Public Property Get DOFEvent(): DOFEvent = m_dof_event: End Property
    Public Property Let DOFEvent(input): m_dof_event = CInt(input): End Property

	Public default Function init()
        m_action = Empty
        m_dof_event = Empty
        Set Init = Me
	End Function

End Class



Class GlfEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "event_player" : End Property

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Raw, newEvent
        
        Dim evtValue, evtValues(), i
        Redim evtValues(UBound(value))
        i=0
        For Each evtValue in value
            Dim newEventValue : Set newEventValue = (new GlfEventDispatch)(evtValue)
            Set evtValues(i) = newEventValue
            i=i+1
        Next
        m_eventValues.Add newEvent.Raw, evtValues  
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            Log "Adding Event Listener for: " & m_events(evt).EventName
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_event_player_play", "EventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        Log "Dispatching Event: " & evt
        If Not IsNull(m_events(evt).Condition) Then
            'msgbox m_events(evt).Condition
            If GetRef(m_events(evt).Condition)(Null) = False Then
                Exit Sub
            End If
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            Log "Dispatching Event: " & evtValue.EventName
            DispatchPinEvent evtValue.EventName, evtValue.Kwargs
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_event_player", message
        End If
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
Function GlfGameSettings()
    Set glf_game = (new GlfGame)()
	Set GlfGameSettings = glf_game
End Function

Class GlfGame

    Private m_balls_per_game

    Public Property Get BallsPerGame()
        BallsPerGame = m_balls_per_game.Value()
    End Property
    Public Property Let BallsPerGame(input)
        Set m_balls_per_game = CreateGlfInput(input)
    End Property

	Public default Function init()

        Set m_balls_per_game = CreateGlfInput(3)
       
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml : yaml = ""
        ToYaml = yaml
    End Function

End Class
Function EnableGlfHighScores()
    Dim high_score_mode : Set high_score_mode = CreateGlfMode("glf_high_scores", 80)
    high_score_mode.StartEvents = Array("game_ending")
    high_score_mode.StopEvents = Array("high_score_complete")
    high_score_mode.UseWaitQueue = True
    Dim high_score : Set high_score = (new GlfHighScore)(high_score_mode)
    high_score.Debug = True
    Set high_score_mode.HighScore = high_score
    Set glf_highscore = high_score
	Set EnableGlfHighScores = glf_highscore
End Function

Class GlfHighScore

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    private m_categories
    private m_defaults
    private m_vars
    private m_award_slide_display_time
    private m_enter_initials_timeout
    private m_reset_high_scores_events
    private m_highscores
    Private m_initials_needed
    Private m_current_initials

    Public Property Get Name() : Name = "high_score" : End Property

    Public Property Get Categories()
        Set Categories = m_categories
    End Property

    Public Property Get Defaults(name)
        Dim new_default : Set new_default = CreateObject("Scripting.Dictionary")
        m_defaults.Add name, new_default
        Set Defaults = m_defaults(name)
    End Property

    Public Property Get Vars(name)
        m_vars.Add name, CreateObject("Scripting.Dictionary")
        Set Vars = m_vars(name)
    End Property

    Public Property Let AwardSlideDisplayTime(input)
        Set m_award_slide_display_time = CreateGlfInput(input)
    End Property

    Public Property Let EnterInitialsTimeout(input)
        Set m_enter_initials_timeout = CreateGlfInput(input)
    End Property

    Public Property Let ResetHighScoreEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_high_scores_events.Add newEvent.Raw, newEvent
        Next
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "high_score_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_award_slide_display_time = CreateGlfInput(4000)
        Set m_enter_initials_timeout = CreateGlfInput(20000)
        Set m_categories = CreateObject("Scripting.Dictionary")
        Set m_defaults = CreateObject("Scripting.Dictionary")
        Set m_vars = CreateObject("Scripting.Dictionary")
        Set m_highscores = CreateObject("Scripting.Dictionary")
        Set m_initials_needed = CreateObject("Scripting.Dictionary")
        Set m_reset_high_scores_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "high_score", Me)
        ReadHighScores()
        Set Init = Me
	End Function

    Public Sub Activate()
        
        'Run High Scores.
        'Check if any player variable beats the stored value for that category.
        'Ask for intiital if we dont have them.
        'display award
        'write scores.
        'finish
        Log "Activating"
        Dim category, p
        
        For Each category in m_categories.Keys()
            'eg score.
            ' get the value for each player
            Dim player_values()
            Dim player_numbers()
            ReDiM player_values(UBound(glf_playerState.Keys()))
            ReDiM player_numbers(UBound(glf_playerState.Keys()))
            For p=0 to UBound(glf_playerState.Keys())
                player_values(p) = GetPlayerStateForPlayer(p, category)
                player_numbers(p) = p+1
            Next
            
             'Sort the values high to low.
             Dim n, i, j, temp, swapped
             n = UBound(player_values) - LBound(player_values) + 1 
             For i = 0 To n - 1
                 swapped = False
                 For j = 0 To n - i - 2 
                     If player_values(j) < player_values(j + 1) Then 
                         temp = player_values(j)
                         player_values(j) = player_values(j + 1)
                         player_values(j + 1) = temp
 
                         temp = player_numbers(j)
                         player_numbers(j) = player_numbers(j + 1)
                         player_numbers(j + 1) = temp
 
                         swapped = True 
                     End If
                 Next
                 If Not swapped Then
                     Exit For
                 End If
             Next

            'msgbox player_values(0)
            
            For i = 0 To UBound(player_values)
                Dim position
                'msgbox "UBound(m_categories(category)): " & UBound(m_categories(category))
                For position=0 to UBound(m_categories(category))
                    Dim high_score
                    'msgbox "Category: " & category
                    high_score = m_highscores(category)(CStr(position+1))("value")
                    high_score = CLng(high_score)
                    ''msgbox "HighScore: " & ">" & high_score & "<"
                    ''msgbox "PlayerScore: " & CInt(player_values(i))
                    'msgbox typename(high_score)
                    'msgbox typename(player_values(i))
                    If player_values(i) > high_score Then
                        'msgbox "setting new high score"
                        Log "Setting new high score"

                        'Shift everything down, knocking the bottom score off.
                        Dim shift_index, hs_item, hs_tmp
                        Set hs_tmp = GlfKwargs()
                        With hs_tmp
                            .Add "label", ""
                            .Add "player_name", ""
                            .Add "value", 0
                        End With
                        Set hs_item = GlfKwargs()
                        With hs_item
                            .Add "label", ""
                            .Add "player_name", ""
                            .Add "value", 0
                        End With
                        For shift_index=position+1 to UBound(m_categories(category))
                            If shift_index>position+1 Then
                                hs_item("label") = hs_tmp("label")
                                hs_item("player_name") = hs_tmp("player_name")
                                hs_item("value") = hs_tmp("value")
                            Else
                                hs_item("label") = m_highscores(category)(CStr(shift_index))("label")
                                hs_item("player_name") = m_highscores(category)(CStr(shift_index))("player_name")
                                hs_item("value") = m_highscores(category)(CStr(shift_index))("value")
                            End If
                            hs_tmp("label") = m_highscores(category)(CStr(shift_index+1))("label")
                            hs_tmp("player_name") = m_highscores(category)(CStr(shift_index+1))("player_name")
                            hs_tmp("value") = m_highscores(category)(CStr(shift_index+1))("value")

                            ''msgbox "hs_item" & hs_item("value")
                            'msgbox "hs_tmp" & hs_tmp("value")

                            m_highscores(category)(CStr(shift_index+1))("label") = hs_item("label")
                            m_highscores(category)(CStr(shift_index+1))("player_name") = hs_item("player_name")
                            m_highscores(category)(CStr(shift_index+1))("value") = hs_item("value")
                        Next

                        'new score
                        m_highscores(category)(CStr(position+1))("value") = player_values(i)
                        m_highscores(category)(CStr(position+1))("player_name") = ""
                        m_highscores(category)(CStr(position+1))("player_num") = player_numbers(i)
                        m_highscores(category)(CStr(position+1))("award") = m_categories(category)(position)
                        m_highscores(category)(CStr(position+1))("category") = category
                        m_highscores(category)(CStr(position+1))("position") = position + 1
                        If Not m_initials_needed.Exists(i+1) Then
                            m_initials_needed.Add i+1, m_highscores(category)(CStr(position+1))
                        End If
                        Exit For
                    Else
                        'msgbox "Player Score " & player_values(i) & " is not greater than the high score: " & high_score
                    End If
                Next
            Next
        Next

        'Ask for Initials
        If Ubound(m_initials_needed.Keys())>-1 Then
            Log "Asking for Initials"
            m_current_initials = 0
            Dim initials_needed_items : initials_needed_items = m_initials_needed.Items()
            AddPinEventListener "text_input_high_score_complete", "text_input_high_score_complete", "HighScoreEventHandler", m_priority, Array("initials_complete", Me, initials_needed_items(0))
            DispatchPinEvent "high_score_enter_initials", initials_needed_items(0)
            SetDelay "enter_initials_timeout", "HighScoreEventHandler", Array(Array("initials_complete", Me, initials_needed_items(0)), Null), m_enter_initials_timeout.Value
        Else
            'No New High Scores, End Mode
            Log "No High Score, Ending"
            'msgbox "No High Score, Ending"
            DispatchPinEvent "high_score_complete", Null
        End If

    End Sub

    Public Sub Deactivate()
        
    End Sub

    Public Sub InitialsInputComplete(initials_item, kwargs)

        'Show Award.
        Log "Initials Complete"
        'msgbox "Initials Complete"
        RemoveDelay "enter_initials_timeout"
        RemovePinEventListener "text_input_high_score_complete", "text_input_high_score_complete"

        'm_highscores(initials_item("category"))(CStr(initials_item("position")))("player_name") = text

        Dim text : text = ""
        If Not IsNull(kwargs) Then
            If kwargs.Exists("text") Then
                text = kwargs("text")
            End If
        End If

        Dim keys, key
        keys = m_highscores.Keys()
        For Each key in keys
            Dim s
            For Each s in m_highscores(key).Keys()
                Dim high_scores_item : Set high_scores_item = m_highscores(key)
                If high_scores_item(s)("player_num") = initials_item("player_num") Then
                    'msgbox "Setting Player " & m_current_initials+1 & " Name to >" & text & "<"
                    high_scores_item(s)("player_name") = text
                End If
            Next
        Next

        Log "Show Award"
        DispatchPinEvent "high_score_award_display", initials_item
        DispatchPinEvent initials_item("award") & "_award_display", initials_item
        DispatchPinEvent initials_item("category") & "_award_display", initials_item
    
        SetDelay "award_display_time", "HighScoreEventHandler", Array(Array("award_display_complete", Me), Null), m_award_slide_display_time.Value

    End Sub

    Public Sub AwardDisplayComplete()

        Log "Award Complete"
        If Ubound(m_initials_needed.Keys())>m_current_initials Then
            Log "Asking for Next Initials"
            m_current_initials = m_current_initials + 1
            Dim initials_needed_items : initials_needed_items = m_initials_needed.Items()
            AddPinEventListener "text_input_high_score_complete", "text_input_high_score_complete", "HighScoreEventHandler", m_priority, Array("initials_complete", Me, initials_needed_items(m_current_initials))
            DispatchPinEvent "high_score_enter_initials", initials_needed_items(m_current_initials)
            SetDelay "enter_initials_timeout", "HighScoreEventHandler", Array(Array("initials_complete", Me, minitials_needed_items(m_current_initials)), Null), m_enter_initials_timeout.Value
        Else
            Log "Writing High Scores"
            Dim keys, key
            keys = m_highscores.Keys()
            Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")
            For Each key in keys
                Dim s, i
                i=1
                'msgbox "Key Count" & ubound(m_highscores(key).Keys())
                For Each s in m_highscores(key).Keys()
                    Dim high_scores_item : high_scores_item = m_highscores(key)
                    'msgbox s
                    tmp.Add key & "_" & i & "_label", m_categories(key)(i-1)
                    tmp.Add key & "_" & i &"_name", high_scores_item(s)("player_name")
                    tmp.Add key & "_" & i &"_value", high_scores_item(s)("value")
                    i=i+1
                Next
                WriteHighScores "HighScores", tmp
            Next
            m_initials_needed.RemoveAll
            Log "Ending"
            DispatchPinEvent "high_score_complete", Null
        End If        
        'End Mode
    End Sub

    Public Sub WriteDefaults()
        ReadHighScores()
        dim key, keys, i
        keys = m_defaults.Keys()
        Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")
        For Each key in keys
            Dim default_keys : default_keys = m_defaults(key).Keys()
            For i=0 to UBound(default_keys)
                Dim default_value_item : Set default_value_item = m_defaults(key)
                Dim cat_a1 : cat_a1 = m_categories(key)
                If m_highscores.Exists(key) Then
                    If Not m_highscores(key).Exists(CStr(i+1)) Then
                        tmp.Add key & "_" & i+1 &"_label", cat_a1(i)
                        tmp.Add key & "_" & i+1 &"_name", default_keys(i)
                        tmp.Add key & "_" & i+1 &"_value", default_value_item(default_keys(i))
                    End If
                Else
                    tmp.Add key & "_" & i+1 & "_label", cat_a1(i)
                    tmp.Add key & "_" & i+1 &"_name", default_keys(i)
                    tmp.Add key & "_" & i+1 &"_value", default_value_item(default_keys(i))
                End If
            Next
            WriteHighScores "HighScores", tmp
        Next  
        ReadHighScores()
    End Sub

    Private Sub WriteHighScores(section, scores)
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
                If scores.Exists(key) Then
                    line = key & "=" & scores(key)
                    scores.Remove key
                End If
            End If
    
            If line = "" And inSection Then
                ' Add remaining keys in the section
                For Each key In scores.Keys
                    outputLines = outputLines & key & "=" & scores(key) & vbCrLf
                Next
                scores.RemoveAll
            End If
            If line <> "" Then
                outputLines = outputLines & line & vbCrLf
            End If
        Next
        
        If Not foundSection Then
            outputLines = outputLines & "["&section&"]" & vbCrLf
            For Each key In scores.Keys
                outputLines = outputLines & key & "=" & scores(key) & vbCrLf
            Next
        End If
        
        Set objFile = objFSO.CreateTextFile(CGameName & "_glf.ini", True)
        objFile.Write outputLines
        objFile.Close
        Glf_ReadMachineVars("HighScores")
    End Sub

    Sub ReadHighScores()
        Dim objFSO, objFile, arrLines, line, inSection
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        If Not objFSO.FileExists(CGameName & "_glf.ini") Then Exit Sub
        
        Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
        arrLines = Split(objFile.ReadAll, vbCrLf)
        objFile.Close
        
        m_highscores.RemoveAll()

        inSection = False
        Dim current_name
        For Each line In arrLines
            line = Trim(line)
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase("HighScores"))
            ElseIf inSection And InStr(line, "=") > 0 Then
                Dim parts : parts = Split(line, "=")
                Dim key_parts : key_parts = Split(parts(0), "_")

                Dim category : category = Trim(key_parts(0))
                Dim position : position = Trim(key_parts(1))
                Dim attr : attr = Trim(key_parts(2))
                Dim current_label
                If m_categories.Exists(category) Then
                    If Not m_highscores.Exists(category) Then
                        m_highscores.Add category, CreateObject("Scripting.Dictionary")
                    End If
                    Select Case attr
                        Case "label":
                            current_label = parts(1)
                        Case "name":
                            current_name = parts(1)
                        Case "value":
                            Dim kwargs : Set kwargs = GlfKwargs()
                            With kwargs
                                .Add "label", current_label
                                .Add "player_name", current_name
                                .Add "value", parts(1)
                            End With
                            'msgbox category & "," & position & ", " & current_label & ", " & current_name & ", " & parts(1)
                            m_highscores(category).Add CStr(position), kwargs
                    End Select
                End If
            End If
        Next
    End Sub
    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_high_score", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml : yaml = ""
        ToYaml = yaml
    End Function

End Class

Function HighScoreEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim high_score : Set high_score = ownProps(1)
    Select Case evt
        Case "activate"
            high_score.Activate
        Case "deactivate"
            high_score.Deactivate
        Case "initials_complete"
            high_score.InitialsInputComplete ownProps(2), kwargs
        Case "award_display_complete"
            high_score.AwardDisplayComplete
    End Select
    If IsObject(args(1)) Then
        Set HighScoreEventHandler = kwargs
    Else
        HighScoreEventHandler = kwargs
    End If
End Function

Class GlfHighScoreItem
	Private m_slide, m_action, m_expire, m_max_queue_time, m_method, m_priority, m_target, m_tokens
    
    Public Property Get Slide(): Slide = m_slide: End Property
    Public Property Let Slide(input)
        m_slide = input
    End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        m_action = input
    End Property

    Public Property Get Expire(): Expire = m_expire: End Property
    Public Property Let Expire(input)
        m_expire = input
    End Property

    Public Property Get MaxQueueTime(): MaxQueueTime = m_max_queue_time: End Property
    Public Property Let MaxQueueTime(input)
        m_max_queue_time = input
    End Property

    Public Property Get Method(): Method = m_method: End Property
    Public Property Let Method(input)
        m_method = input
    End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input)
        m_priority = input
    End Property

    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input)
        m_target = input
    End Property

    Public Property Get Tokens(): Tokens = m_tokens: End Property
    Public Property Let Tokens(input)
        m_tokens = input
    End Property

	Public default Function init()
        m_action = "play"
        m_slide = Empty
        m_expire = Empty
        m_max_queue_time = Empty
        m_method = Empty
        m_priority = Empty
        m_target = Empty
        m_tokens = Empty
        Set Init = Me
	End Function

End Class


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
    Public Property Get EventName(name)
        If m_events.Exists(name) Then
            Set EventName = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfLightPlayerEventItem)()
            m_events.Add name, new_event
            Set EventName = new_event
        End If
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
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
            Log "Adding Event Listener for event: " & evt
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", m_priority, Array("play", Me, m_events(evt), evt)
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
        Log "Reloading Lights"
        Dim evt
        For Each evt in m_events.Keys()
            Dim lightName, light,lightsCount,x,tagLight, tagLights
            'First get light counts
            lightsCount = 0
            For Each lightName in m_events(evt).LightNames
                Set light = m_events(evt).Lights(lightName)
                If Not glf_lightNames.Exists(lightName) Then
                    tagLights = glf_lightTags("T_"&lightName).Keys()
                    Log "Tag Lights: " & Join(tagLights)
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            Next
            Log "Adding " & lightsCount & " lights for event: " & evt 
            Dim seqArray
            ReDim seqArray(lightsCount-1)
            x=0
            'Build Seq
            For Each lightName in m_events(evt).LightNames
                Set light = m_events(evt).Lights(lightName)

                If Not glf_lightNames.Exists(lightName) Then
                    tagLights = glf_lightTags("T_"&lightName).Keys()
                    For Each tagLight in tagLights
                        seqArray(x) = tagLight & "|100|" & light.Color & "|" & light.Fade
                        x=x+1
                    Next
                Else
                    seqArray(x) = lightName & "|100|" & light.Color & "|" & light.Fade
                    x=x+1
                End If
            Next
            Log "Light List: " & Join(seqArray)
            m_events(evt).LightSeq = seqArray
        Next   
    End Sub

    Public Sub Play(evt, lights)
        LightPlayerCallbackHandler evt, Array(lights.LightSeq), m_name, m_priority, True, 1, Empty
    End Sub

    Public Sub PlayOff(evt, lights)
        LightPlayerCallbackHandler evt, Array(lights.LightSeq), m_name, m_priority, False, 1, Empty
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt, key
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
        Dim yaml, key
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

Function LightPlayerCallbackHandler(key, lights, mode, priority, play, speed, color_replacement)
    Dim shows_added
    Dim lightStack
    Dim lightParts, light
    If play = False Then
        For Each light in lights(0)
            lightParts = Split(light,"|")
            Set lightStack = glf_lightStacks(lightParts(0))
            If Not lightStack.IsEmpty() Then
                'glf_debugLog.WriteToLog "LightPlayer", "Removing Light " & lightParts(0)
                lightStack.PopByKey(mode & "_" & key)
                Dim show_key
                For Each show_key in glf_running_shows.Keys
                    If Left(show_key, Len("fade_" & mode & "_" & key & "_" & lightParts(0))) = "fade_" & mode & "_" & key & "_" & lightParts(0) Then
                        glf_running_shows(show_key).StopRunningShow()
                    End If
                Next
                If Not lightStack.IsEmpty() Then
                    ' Set the light to the next color on the stack
                    Dim nextColor
                    Set nextColor = lightStack.Peek()
                    Glf_SetLight lightParts(0), nextColor("Color")
                Else
                    ' Turn off the light since there's nothing on the stack
                    Glf_SetLight lightParts(0), "000000"
                End If
            End If
        Next
        LightPlayerCallbackHandler = Array(Null)
        Exit Function
        'glf_debugLog.WriteToLog "LightPlayer", "Removing Light Seq" & mode & "_" & key
    Else
        If UBound(lights) = -1 Then
            LightPlayerCallbackHandler = Array(Null)
            Exit Function
        End If
        If IsArray(lights) Then
            'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key    
        Else
            'glf_debugLog.WriteToLog "LightPlayer", "Lights not an array!?"
        End If
        'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key
        Set shows_added = CreateObject("Scripting.Dictionary")
        For Each light in lights(0)
            lightParts = Split(light,"|")
            
            Set lightStack = glf_lightStacks(lightParts(0))
            Dim oldColor : oldColor = Empty
            Dim newColor : newColor = lightParts(2)
            If Not IsEmpty(color_replacement) Then
                newColor = color_replacement
            End If

            If lightStack.IsEmpty() Then
                oldColor = "000000"
                ' If stack is empty, push the color onto the stack and set the light color
                lightStack.Push mode & "_" & key, newColor, priority
                Glf_SetLight lightParts(0), newColor
            Else
                Dim current
                Set current = lightStack.Peek()                
                If priority >= current("Priority") Then
                    oldColor = current("Color")
                    ' If the new priority is higher, push it onto the stack and change the light color
                    lightStack.Push mode & "_" & key, newColor, priority
                    Glf_SetLight lightParts(0), newColor
                Else
                    ' Otherwise, just push it onto the stack without changing the light color
                    lightStack.Push mode & "_" & key, newColor, priority
                End If
            End If
            
            If Not IsEmpty(oldColor) And Ubound(lightParts)=4 Then
                If lightParts(4) <> "" Then
                    'FadeMs
                    Dim cache_name, new_running_show,cached_show,show_settings, fade_seq
                    cache_name = lightParts(3)
                    fade_seq = Glf_FadeRGB(oldColor, newColor, lightParts(4))
                    'MsgBox cache_name
                    shows_added.Add cache_name, True
                    
                    If glf_cached_shows.Exists(cache_name & "__-1") Then
                        Set show_settings = (new GlfShowPlayerItem)()
                        show_settings.Show = cache_name
                        show_settings.Loops = 1
                        show_settings.Speed = speed
                        show_settings.ColorLookup = fade_seq
                        Set new_running_show = (new GlfRunningShow)(cache_name, show_settings.Key, show_settings, priority+1, Null, Null)
                    End If
                End If
            End If
        Next
        LightPlayerCallbackHandler = Array(shows_added)
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

    Public Function PopByKey(key)
        Dim i, removedItem, found
        found = False
        Set removedItem = Nothing
    
        ' Loop through the stack to find the item with the matching key
        For i = LBound(stack) To UBound(stack)
            If stack(i)("Key") = key Then
                ' Store the item to be removed
                Set removedItem = stack(i)
                found = True
    
                ' Shift all elements after the removed item to the left
                Dim j
                For j = i To UBound(stack) - 1
                    Set stack(j) = stack(j + 1)
                Next
    
                ' Resize the array to remove the last element
                ReDim Preserve stack(UBound(stack) - 1)
                Exit For
            End If
        Next
    
        ' Return the removed item (or Nothing if not found)
        If found Then
            Set PopByKey = removedItem
        Else
            Set PopByKey = Nothing
        End If
    End Function
    

    ' Get the current top color without popping it
    Public Function Peek()
        If UBound(stack) >= 0 Then
            Set Peek = stack(LBound(stack))
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

    Public Sub PrintStackOrder()
        Dim i
        Debug.Print "Stack Order:" 
        For i = LBound(stack) To UBound(stack)
            Debug.Print "Key: " & stack(i)("Key") & ", Color: " & stack(i)("Color") & ", Priority: " & stack(i)("Priority")
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
            m_enable_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Get DisableEvents(): Set DisableEvents = m_disable_events: End Property
    Public Property Let DisableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_events.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Get Mode(): Set Mode = m_mode: End Property

    Public Property Let Debug(value) : m_debug = value : End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode, device, parent)
        Set m_mode = mode
        m_priority = mode.Priority
        m_device = device
        Set m_parent = parent
        If mode.IsDebug = 1 Then
            m_debug = True
        End If

        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_disable_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode.Name & "_starting", m_device & "_" & m_parent.Name & "_activate", "BaseModeDeviceEventHandler", m_priority+1, Array("activate", Me)
        AddPinEventListener m_mode.Name & "_stopping", m_device & "_" & m_parent.Name & "_deactivate", "BaseModeDeviceEventHandler", m_priority-1, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_enable_events.Keys()
            AddPinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable", "BaseModeDeviceEventHandler", m_priority+m_enable_events(evt).Priority, Array("enable", m_parent, m_enable_events(evt))
        Next
        For Each evt In m_disable_events.Keys()
            AddPinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable", "BaseModeDeviceEventHandler", m_priority+m_disable_events(evt).Priority, Array("disable", m_parent, m_disable_events(evt))
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
            glf_debugLog.WriteToLog m_mode.Name & "_" & m_device & "_" & m_parent.Name, message
        End If
    End Sub
End Class


Function BaseModeDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
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
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            device.Enable
        Case "disable"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            device.Disable
    End Select
    If IsObject(args(1)) Then
        Set BaseModeDeviceEventHandler = kwargs
    Else
        BaseModeDeviceEventHandler = kwargs
    End If
End Function

Class Mode

    Private m_name
    Private m_modename 
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
    Private m_queueEventplayer
    Private m_queueRelayPlayer
    Private m_random_event_player
    Private m_sound_player
    Private m_dof_player
    Private m_slide_player
    Private m_shot_profiles
    Private m_sequence_shots
    Private m_state_machines
    Private m_extra_balls
    Private m_combo_switches
    Private m_timed_switches
    Private m_tilt
    Private m_high_score
    Private m_use_wait_queue

    Public Property Get Name(): Name = m_name : End Property
    Public Property Get ModeName(): ModeName = m_modename : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "active":
                If m_started Then
                    GetValue = True
                Else
                    GetValue = False
                End If
        End Select
    End Property
    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Get Status()
        If m_started Then
            Status = "started"
        Else
            Status = "stopped"
        End If
    End Property
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
    Public Property Get QueueEventPlayer() : Set QueueEventPlayer = m_queueEventplayer: End Property
    Public Property Get QueueRelayPlayer() : Set QueueRelayPlayer = m_queueRelayPlayer: End Property
    Public Property Get RandomEventPlayer() : Set RandomEventPlayer = m_random_event_player : End Property
    Public Property Get VariablePlayer(): Set VariablePlayer = m_variableplayer: End Property
    Public Property Get SoundPlayer() : Set SoundPlayer = m_sound_player : End Property
    Public Property Get DOFPlayer() : Set DOFPlayer = m_dof_player : End Property
    Public Property Get SlidePlayer() : Set SlidePlayer = m_slide_player : End Property

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

    Public Property Get BallSavesItems() : BallSavesItems = m_ballsaves.Items() : End Property
    Public Property Get BallSaves(name)
        If m_ballsaves.Exists(name) Then
            Set BallSaves = m_ballsaves(name)
        Else
            Dim new_ballsave : Set new_ballsave = (new BallSave)(name, Me)
            m_ballsaves.Add name, new_ballsave
            Set BallSaves = new_ballsave
        End If
    End Property

    Public Property Get TimersItems() : TimersItems = m_timers.Items() : End Property
    Public Property Get Timers(name)
        If m_timers.Exists(name) Then
            Set Timers = m_timers(name)
        Else
            Dim new_timer : Set new_timer = (new GlfTimer)(name, Me)
            m_timers.Add name, new_timer
            Set Timers = new_timer
        End If
    End Property

    Public Property Get CountersItems() : CountersItems = m_counters.Items() : End Property
    Public Property Get Counters(name)
        If m_counters.Exists(name) Then
            Set Counters = m_counters(name)
        Else
            Dim new_counter : Set new_counter = (new GlfCounter)(name, Me)
            m_counters.Add name, new_counter
            Set Counters = new_counter
        End If
    End Property

    Public Property Get MultiballLocksItems() : MultiballLocksItems = m_multiball_locks.Items() : End Property
    Public Property Get MultiballLocks(name)
        If m_multiball_locks.Exists(name) Then
            Set MultiballLocks = m_multiball_locks(name)
        Else
            Dim new_multiball_lock : Set new_multiball_lock = (new GlfMultiballLocks)(name, Me)
            m_multiball_locks.Add name, new_multiball_lock
            Set MultiballLocks = new_multiball_lock
        End If
    End Property

    Public Property Get MultiballsItems() : MultiballsItems = m_multiballs.Items() : End Property
    Public Property Get Multiballs(name)
        If m_multiballs.Exists(name) Then
            Set Multiballs = m_multiballs(name)
        Else
            Dim new_multiball : Set new_multiball = (new GlfMultiballs)(name, Me)
            m_multiballs.Add name, new_multiball
            Set Multiballs = new_multiball
        End If
    End Property

    Public Property Get SequenceShotsItems() : SequenceShotsItems = m_sequence_shots.Items() : End Property
    Public Property Get SequenceShots(name)
        If m_sequence_shots.Exists(name) Then
            Set SequenceShots = m_sequence_shots(name)
        Else
            Dim new_sequence_shot : Set new_sequence_shot = (new GlfSequenceShots)(name, Me)
            m_sequence_shots.Add name, new_sequence_shot
            Set SequenceShots = new_sequence_shot
        End If
    End Property

    Public Property Get ExtraBallsItems() : ExtraBallsItems = m_extra_balls.Items() : End Property
    Public Property Get ExtraBalls(name)
        If m_extra_balls.Exists(name) Then
            Set ExtraBalls = m_extra_balls(name)
        Else
            Dim new_extra_ball : Set new_extra_ball = (new GlfExtraBall)(name, Me)
            m_extra_balls.Add name, new_extra_ball
            Set ExtraBalls = new_extra_ball
        End If
    End Property
    Public Property Get ComboSwitchesItems() : ComboSwitchesItems = m_combo_switches.Items() : End Property
    Public Property Get ComboSwitches(name)
        If m_combo_switches.Exists(name) Then
            Set ComboSwitches = m_combo_switches(name)
        Else
            Dim new_combo_switch : Set new_combo_switch = (new GlfComboSwitches)(name, Me)
            m_combo_switches.Add name, new_combo_switch
            Set ComboSwitches = new_combo_switch
        End If
    End Property
    Public Property Get TimedSwitchesItems() : TimedSwitchesItems = m_timed_switches.Items() : End Property
        Public Property Get TimedSwitches(name)
            If m_timed_switches.Exists(name) Then
                Set TimedSwitches = m_timed_switches(name)
            Else
                Dim new_timed_switch : Set new_timed_switch = (new GlfTimedSwitches)(name, Me)
                m_timed_switches.Add name, new_timed_switch
                Set TimedSwitches = new_timed_switch
            End If
        End Property
    Public Property Get Tilt()
        If Not IsNull(m_tilt) Then
            Set Tilt = m_tilt
        Else
            Set m_tilt = (new GlfTilt)(Me)
            Set Tilt = m_tilt
        End If
    End Property

    Public Property Get TiltConfig()
        If Not IsNull(m_tilt) Then
            Set TiltConfig = m_tilt
        Else
            TiltConfig = Null
        End If
    End Property

    Public Property Set HighScore(value) : Set m_high_score = value : End Property
    Public Property Get HighScore()
        If Not IsNull(m_high_score) Then
            Set HighScore = m_high_score
        Else
            Set m_high_score = (new GlfHighScore)(Me)
            Set HighScore = m_high_score
        End If
    End Property

    Public Property Get StateMachines(name)
        If m_state_machines.Exists(name) Then
            Set StateMachines = m_state_machines(name)
        Else
            Dim new_state_machine : Set new_state_machine = (new GlfStateMachine)(name, Me)
            m_state_machines.Add name, new_state_machine
            Set StateMachines = new_state_machine
        End If
    End Property
    Public Property Get ModeStateMachines(): ModeStateMachines = m_state_machines.Items(): End Property

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

    Public Property Get ShotGroupsItems() : ShotGroupsItems = m_shot_groups.Items() : End Property
    Public Property Get ShotGroups(name)
        If m_shot_groups.Exists(name) Then
            Set ShotGroups = m_shot_groups(name)
        Else
            Dim new_shot_group : Set new_shot_group = (new GlfShotGroup)(name, Me)
            m_shot_groups.Add name, new_shot_group
            Set ShotGroups = new_shot_group
        End If
    End Property

    Public Property Get BallHoldsItems() : BallHoldsItems = m_ballholds.Items() : End Property
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
            m_start_events.Add newEvent.Raw, newEvent
            AddPinEventListener newEvent.EventName, m_name & "_start", "ModeEventHandler", m_priority+newEvent.Priority, Array("start", Me, newEvent)
        Next
    End Property
    
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Raw, newEvent
            AddPinEventListener newEvent.EventName, m_name & "_stop", "ModeEventHandler", m_priority+newEvent.Priority+1, Array("stop", Me, newEvent)
        Next
    End Property

    Public Property Get UseWaitQueue(): UseWaitQueue = m_use_wait_queue: End Property
    Public Property Let UseWaitQueue(input): m_use_wait_queue = input: End Property

    Public Property Get IsDebug()
        If m_debug = True Then
            IsDebug = 1
        Else
            IsDebug = 0
        End If
    End Property
    Public Property Let Debug(value)
        m_debug = value
        Dim config_item
        For Each config_item in m_ballsaves.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_counters.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_timers.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_multiball_locks.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_multiballs.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_shots.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_shot_groups.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_ballholds.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_sequence_shots.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_state_machines.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_extra_balls.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_combo_switches.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_timed_switches.Items()
            config_item.Debug = value
        Next
        If Not IsNull(m_tilt) Then
            m_tilt.Debug = value
        End If
        If Not IsNull(m_lightplayer) Then
            m_lightplayer.Debug = value
        End If
        If Not IsNull(m_eventplayer) Then
            m_eventplayer.Debug = value
        End If
        If Not IsNull(m_queueEventplayer) Then
            m_queueEventplayer.Debug = value
        End If
        If Not IsNull(m_queueRelayPlayer) Then
            m_queueRelayPlayer.Debug = value
        End If
        If Not IsNull(m_random_event_player) Then
            m_random_event_player.Debug = value
        End If
        If Not IsNull(m_sound_player) Then
            m_sound_player.Debug = value
        End If
        If Not IsNull(m_dof_player) Then
            m_dof_player.Debug = value
        End If
        If Not IsNull(m_slide_player) Then
            m_slide_player.Debug = value
        End If
        If Not IsNull(m_showplayer) Then
            m_showplayer.Debug = value
        End If
        If Not IsNull(m_segment_display_player) Then
            m_segment_display_player.Debug = value
        End If
        If Not IsNull(m_variableplayer) Then
            m_variableplayer.Debug = value
        End If
        Glf_MonitorModeUpdate Me
    End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_modename = name
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
        Set m_sequence_shots = CreateObject("Scripting.Dictionary")
        Set m_state_machines = CreateObject("Scripting.Dictionary")
        Set m_extra_balls = CreateObject("Scripting.Dictionary")
        Set m_combo_switches = CreateObject("Scripting.Dictionary")
        Set m_timed_switches = CreateObject("Scripting.Dictionary")

        m_use_wait_queue = False
        m_lightplayer = Null
        m_tilt = Null
        m_showplayer = Null
        m_segment_display_player = Null
        m_high_score = Null
        Set m_eventplayer = (new GlfEventPlayer)(Me)
        Set m_queueEventplayer = (new GlfQueueEventPlayer)(Me)
        Set m_queueRelayPlayer = (new GlfQueueRelayPlayer)(Me)
        Set m_random_event_player = (new GlfRandomEventPlayer)(Me)
        Set m_sound_player = (new GlfSoundPlayer)(Me)
        Set m_dof_player = (new GlfDofPlayer)(Me)
        Set m_slide_player = (new GlfSlidePlayer)(Me)
        Set m_variableplayer = (new GlfVariablePlayer)(Me)
        Glf_MonitorModeUpdate Me
        AddPinEventListener m_name & "_starting", m_name & "_starting_end", "ModeEventHandler", -99, Array("started", Me, "")
        AddPinEventListener m_name & "_stopping", m_name & "_stopping_end", "ModeEventHandler", -99, Array("stopped", Me, "")
        Set Init = Me
	End Function

    Public Sub StartMode()
        Log "Starting"
        m_started=True
        If useBcp Then
            bcpController.ModeStart m_modename, m_priority
        End If
        DispatchQueuePinEvent m_name & "_starting", Null
    End Sub

    Public Sub StopMode()
        If m_started = True Then
            m_started = False
            Log "Stopping"
            If useBcp Then
                bcpController.SlidesClear(m_modename)
                bcpController.ModeStop(m_modename)
            End If
            DispatchQueuePinEvent m_name & "_stopping", Null
        End If
    End Sub

    Public Sub Started()
        DispatchPinEvent m_name & "_started", Null
        Glf_MonitorModeUpdate Me
        glf_running_modes = glf_running_modes & "["""&m_modename&""", " & m_priority & "],"
        Log "Started"
    End Sub

    Public Sub Stopped()
        'MsgBox m_name & "Stopped"
        DispatchPinEvent m_name & "_stopped", Null
        Glf_MonitorModeUpdate Me
        glf_running_modes = Replace(glf_running_modes, "["""&m_modename&""", " & m_priority & "],", "")
        Log "Stopped"
    End Sub

    Private Sub Log(message) 
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        dim yaml, child,x, key

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

        If UBound(m_combo_switches.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "combo_switches: " & vbCrLf
            For Each child in m_combo_switches.Keys
                yaml = yaml & m_combo_switches(child).ToYaml
            Next
        End If

        If UBound(m_sequence_shots.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "sequence_shots: " & vbCrLf
            For Each child in m_sequence_shots.Keys
                yaml = yaml & m_sequence_shots(child).ToYaml
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
                yaml = yaml & m_lightplayer.ToYaml()
            End If
        End If

        If Not IsNull(m_slide_player) Then
            If UBound(m_slide_player.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "slide_player: " & vbCrLf
                yaml = yaml & m_slide_player.ToYaml()
            End If
        End If

        If Not IsNull(m_variableplayer) Then
            If UBound(m_variableplayer.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "variable_player: " & vbCrLf
                yaml = yaml & m_variableplayer.ToYaml()
            End If
        End If

        If Not IsNull(m_random_event_player) Then
            If UBound(m_random_event_player.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "random_event_player: " & vbCrLf
                yaml = yaml & m_random_event_player.ToYaml()
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
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            mode.StartMode
            If mode.UseWaitQueue = True Then
                kwargs.Add "wait_for", mode.Name & "_stopped"
            End If
        Case "stop"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            mode.StopMode
        Case "started"
            mode.Started
        Case "stopped"
            mode.Stopped
    End Select
    If IsObject(args(1)) Then
        Set ModeEventHandler = kwargs
    Else
        ModeEventHandler = kwargs
    End If
End Function

Class GlfMultiballLocks

    Private m_name
    Private m_lock_devices
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_enable_events
    Private m_disable_events
    Private m_balls_to_lock
    Private m_balls_locked
    Private m_balls_to_replace
    Private m_lock_events
    Private m_reset_events
    Private m_enabled
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "locked_balls":
                GetValue = m_balls_locked
        End Select
    End Property
    Public Property Get LockDevices() : LockDevices = m_lock_devices : End Property
    Public Property Let LockDevices(value) : m_lock_devices = value : End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let BallsToLock(value) : m_balls_to_lock = value : End Property
    Public Property Let LockEvents(value) : m_lock_events = value : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let BallsToReplace(value) : m_balls_to_replace = value : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(name, mode)
        m_name = "multiball_lock_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_lock_events = Array()
        m_reset_events = Array()
        m_lock_devices = Array()
        m_balls_to_lock = 0
        m_balls_to_replace = -1
        m_enabled = False
        m_balls_locked = 0
        Set m_base_device = (new GlfBaseModeDevice)(mode, "multiball_lock", Me)
        glf_multiball_locks.Add name, Me
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
        Log "Enabling"
        m_enabled = True
        Dim lock_device
        For Each lock_device in m_lock_devices
            AddPinEventListener "balldevice_" & lock_device & "_ball_enter", m_mode & "_" & name & "_lock", "MultiballLocksHandler", m_priority, Array("lock", me, lock_device)
        Next
        Dim evt
        For Each evt in m_lock_events
            AddPinEventListener evt, m_name & "_ball_locked", "MultiballLocksHandler", m_priority, Array("virtual_lock", Me, Null)
        Next
        For Each evt in m_reset_events
            AddPinEventListener evt, m_name & "_reset", "MultiballLocksHandler", m_priority, Array("reset", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Dim lock_device
        For Each lock_device in m_lock_devices
            RemovePinEventListener "balldevice_" & lock_device & "_ball_enter", m_mode & "_" & name & "_lock"
        Next
        Dim evt
        For Each evt in m_lock_events
            RemovePinEventListener evt, m_name & "_ball_locked"
        Next
        For Each evt in m_reset_events
            RemovePinEventListener evt, m_name & "_reset"
        Next
    End Sub

    Public Function Lock(device, unclaimed_balls)
        Log "Locking for device: " & device & ". Unclaimed Balls: " & unclaimed_balls
        If unclaimed_balls <= 0 Then
            Lock = unclaimed_balls
            Exit Function
        End If
        
        Dim balls_locked
        If GetPlayerState(m_name & "_balls_locked") = False Then
            balls_locked = 1
        Else
            balls_locked = GetPlayerState(m_name & "_balls_locked") + 1
        End If
        If balls_locked > m_balls_to_lock Then
            Log "Cannot lock balls. Lock is full."
            Lock = unclaimed_balls
            Exit Function
        End If

        SetPlayerState m_name & "_balls_locked", balls_locked
        

        If Not IsNull(device) Then
            
            If glf_ball_devices(device).Balls() > balls_locked Then
                glf_ball_devices(device).Eject()
            Else
                If m_balls_to_replace = -1 Or balls_locked <= m_balls_to_replace Then
                    ' glf_BIP = glf_BIP - 1
                    SetDelay m_name & "_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", Me),Null), 1000
                End If
            End If
        End If

        DispatchPinEvent m_name & "_locked_ball", balls_locked
        
        If balls_locked = m_balls_to_lock Then
            DispatchPinEvent m_name & "_full", balls_locked
        End If

        Lock = unclaimed_balls - 1
    End Function

    Public Sub Reset
        Log "Resetting multiball lock count"
        SetPlayerState m_name & "_balls_locked", 0
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function MultiballLocksHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
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
            kwargs = multiball.Lock(ownProps(2), kwargs)
        Case "virtual_lock"
            multiball.Lock Null, 1
        Case "reset"
            multiball.Reset
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True And glf_plunger.IncomingBalls = 0  Then
                Glf_ReleaseBall(Null)
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
    If IsObject(args(1)) Then
        Set MultiballLocksHandler = kwargs
    Else
        MultiballLocksHandler = kwargs
    End If
End Function
Class GlfMultiballs

    Private m_name
    Private m_configname
    Private m_mode
    Private m_priority
    Private m_base_device
    Private m_ball_count
    Private m_ball_locks
    Private m_add_a_ball_events
    Private m_add_a_ball_grace_period
    Private m_add_a_ball_hurry_up_time
    Private m_add_a_ball_shoot_again
    Private m_ball_count_type
    Private m_disable_events
    Private m_enable_events
    Private m_grace_period
    Private m_grace_period_enabled
    Private m_hurry_up
    Private m_hurry_up_enabled
    Private m_replace_balls_in_play
    Private m_reset_events
    Private m_shoot_again
    Private m_source_playfield
    Private m_start_events
    Private m_stop_events
    Private m_balls_added_live
    Private m_balls_live_target
    Private m_enabled
    Private m_shoot_again_enabled
    Private m_queued_balls
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
        End Select
    End Property

    Public Property Let BallCount(value): Set m_ball_count = CreateGlfInput(value): End Property
    Public Property Let AddABallEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_add_a_ball_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let AddABallGracePeriod(value): Set m_add_a_ball_grace_period = CreateGlfInput(value): End Property
    Public Property Let AddABallHurryUpTime(value): Set m_add_a_ball_hurry_up_time = CreateGlfInput(value): End Property
    Public Property Let AddABallShootAgain(value): Set m_add_a_ball_shoot_again = CreateGlfInput(value): End Property
    Public Property Let BallCountType(value): m_ball_count_type = value: End Property
    Public Property Let BallLocks(value): m_ball_locks = value: End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let GracePeriod(value): Set m_grace_period = CreateGlfInput(value): End Property
    Public Property Let HurryUp(value): Set m_hurry_up = CreateGlfInput(value): End Property
    Public Property Let ReplaceBallsInPlay(value): m_replace_balls_in_play = value: End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ShootAgain(value): Set m_shoot_again = CreateGlfInput(value): End Property
    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Raw, newEvent
        Next
    End Property
        
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(name, mode)
        m_name = "multiball_" & name
        m_configname = name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_ball_count = CreateGlfInput(0)
        Set m_add_a_ball_events = CreateObject("Scripting.Dictionary")
        Set m_add_a_ball_grace_period = CreateGlfInput(0)
        Set m_add_a_ball_hurry_up_time = CreateGlfInput(0)
        Set m_add_a_ball_shoot_again = CreateGlfInput(5000)
        m_ball_count_type = "total"
        m_ball_locks = Array()
        Set m_grace_period = CreateGlfInput(0)
        Set m_hurry_up = CreateGlfInput(0)
        m_replace_balls_in_play = False
        Set m_shoot_again = CreateGlfInput(10000)
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")
        m_replace_balls_in_play = False
        m_balls_added_live = 0
        m_balls_live_target = 0
        m_queued_balls = 0
        m_enabled = False
        m_shoot_again_enabled = False
        m_grace_period_enabled = False
        m_hurry_up_enabled = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "multiball", Me)
        glf_multiballs.Add name, Me
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
        Log "Enabling " & m_name
        m_enabled = True
        Dim evt
        For Each evt in m_start_events.Keys
            AddPinEventListener m_start_events(evt).EventName, m_name & "_" & evt & "_start", "MultiballsHandler", m_priority+m_start_events(evt).Priority, Array("start", Me, m_start_events(evt))
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset", "MultiballsHandler", m_priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_add_a_ball_events.Keys
            AddPinEventListener m_add_a_ball_events(evt).EventName, m_name & "_" & evt & "_add_a_ball", "MultiballsHandler", m_priority, Array("add_a_ball", Me, m_add_a_ball_events(evt))
        Next
        For Each evt in m_stop_events.Keys
            AddPinEventListener m_stop_events(evt).EventName, m_name & "_" & evt & "_stop", "MultiballsHandler", m_priority+m_stop_events(evt).Priority, Array("stop", Me, m_stop_events(evt))
        Next
    End Sub
    
    Public Sub Disable()
        Log "Disabling " & m_name
        m_enabled = False
        m_balls_added_live = 0
        m_balls_live_target = 0
        m_shoot_again_enabled = False
        StopMultiball()
        Dim evt
        For Each evt in m_start_events.Keys
            RemovePinEventListener m_start_events(evt).EventName, m_name & "_" & evt & "_start"
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_add_a_ball_events.Keys
            RemovePinEventListener m_add_a_ball_events(evt).EventName, m_name & "_" & evt & "_add_a_ball"
        Next
        For Each evt in m_stop_events.Keys
            RemovePinEventListener m_stop_events(evt).EventName, m_name & "_" & evt & "_stop"
        Next
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        'RemoveDelay m_name & "_queued_release"
    End Sub
    
    Private Sub HandleBallsInPlayAndBallsLive()
        'Dim balls_to_replace
        'If m_replace_balls_in_play = True Then
        '    balls_to_replace = glf_BIP
        'Else
        '    balls_to_replace = 0
        'End If
        'Log("Going to add an additional " & balls_to_replace & " balls for replace_balls_in_play")
        m_balls_added_live = 0 
        Dim ball_count_value : ball_count_value = m_ball_count.Value
        If m_ball_count_type = "total" Then
            Log "glf_BIP: " & glf_BIP
            If ball_count_value > glf_BIP Then
                m_balls_added_live = ball_count_value - glf_BIP
                'glf_BIP = m_ball_count
            End If
            m_balls_live_target = ball_count_value
        Else
            m_balls_added_live = ball_count_value
            'glf_BIP = glf_BIP + m_balls_added_live
            m_balls_live_target = glf_BIP + m_balls_added_live
        End If

    End Sub

    Public Function BallsDrained(balls)
        If m_shoot_again_enabled Then
            balls = BallDrainShootAgain(balls)
        Else
            BallDrainCountBalls(balls)
        End If
        BallsDrained = balls
    End Function

    Public Sub Start()
        ' Start multiball.
        If not m_enabled Then
            Exit Sub
        End If

        If m_balls_live_target > 0 Then
            Log("Cannot start MB because " & m_balls_live_target & " are still in play")
            Exit Sub
        End If

        m_shoot_again_enabled = True

        HandleBallsInPlayAndBallsLive()
        Log("Starting multiball with " & m_balls_live_target & " balls (added " & m_balls_added_live & ")")
        'msgbox("Starting multiball with " & m_balls_live_target & " balls (added " & m_balls_added_live & ")")    
        Dim balls_added : balls_added = 0

        'eject balls from locks
        Dim ball_lock
        For Each ball_lock in m_ball_locks
            Dim available_balls : available_balls = glf_ball_devices(ball_lock).Balls()
            If available_balls > 0 Then
                glf_ball_devices(ball_lock).EjectAll()
            End If
            balls_added = available_balls
        Next

        glf_BIP = m_balls_live_target

        'request remaining balls
        m_queued_balls = (m_balls_added_live - balls_added)
        If m_queued_balls > 0 Then
            SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me),Null), 1000
        End If

        If m_shoot_again.Value = 0 Then
            'No shoot again. Just stop multiball right away
            StopMultiball()
        else
            'Enable shoot again
            TimerStart()
        End If
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "MultiballsHandler", m_priority, Array("drain", Me)

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls", m_balls_live_target
        End With
        DispatchPinEvent m_name & "_started", kwargs
    End Sub

    Sub TimerStart()
        DispatchPinEvent "ball_save_" & m_configname & "_timer_start", Null 'desc: The multiball ball save called (name) has just start its countdown timer.
        StartShootAgain m_shoot_again.Value, m_grace_period.Value, m_hurry_up.Value
    End Sub

    Sub StartShootAgain(shoot_again_ms, grace_period_ms, hurry_up_time_ms)
        'Set callbacks for shoot again, grace period, and hurry up, if values above 0 are provided.
        'This is started for both beginning multiball ball save and add a ball ball save
        If shoot_again_ms > 0 Then
            Log("Starting ball save timer: " & shoot_again_ms)
            SetDelay m_name&"_disable_shoot_again", "MultiballsHandler" , Array(Array("stop", Me),Null), shoot_again_ms+grace_period_ms
        End If
        If grace_period_ms > 0 Then
            m_grace_period_enabled = True
            SetDelay m_name&"_grace_period", "MultiballsHandler" , Array(Array("grace_period", Me),Null), shoot_again_ms
        End If
        If hurry_up_time_ms > 0 Then
            m_hurry_up_enabled = True
            SetDelay m_name&"_hurry_up", "MultiballsHandler" , Array(Array("hurry_up", Me),Null), shoot_again_ms - hurry_up_time_ms
        End If
    End Sub

    Sub RunHurryUp()
        Log("Starting Hurry Up")
        m_hurry_up_enabled = False
        DispatchPinEvent m_name & "_hurry_up", Null
    End Sub

    Sub RunGracePeriod()
        Log("Starting Grace Period")
        m_grace_period_enabled = False
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Function BallDrainShootAgain(balls):
        Dim balls_to_save, kwargs

        If balls = 0 Then
            BallDrainShootAgain = balls
            Exit Function
        End If

        balls_to_save = m_balls_live_target - balls

        Log "Balls to save: " & balls_to_save & ". Balls live target: " & m_balls_live_target & ". Balls in Play: " & glf_BIP & ". Balls Drained: " & balls
        
        If balls_to_save <= 0 Then
            BallDrainShootAgain = balls
        End If

        If balls_to_save > balls Then
            balls_to_save = balls
        End If

        Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls", balls_to_save
        End With
        DispatchPinEvent m_name & "_shoot_again", kwargs
        
        Log("Ball drained during MB. Requesting a new one")
        m_queued_balls = m_queued_balls + 1
        SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me, m_queued_balls),Null), 1000

        BallDrainShootAgain = balls - balls_to_save
    End Function

    Function BallDrainCountBalls(balls):
        DispatchPinEvent m_name & "_ball_lost", Null
        If not glf_gameStarted or (glf_BIP - balls) = 1 Then
            m_balls_added_live = 0
            m_balls_live_target = 0
            DispatchPinEvent m_name & "_ended", Null
            RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
            Log("Ball drained. MB ended.")
        End If
        BallDrainCountBalls = balls
    End Function

    Public Sub Reset()
        Log "Resetting multiball: " & m_name
        DispatchPinEvent m_name & "_reset_event", Null

        Disable()
        m_shoot_again_enabled = False
        m_balls_added_live = 0
        m_balls_live_target = 0
    End Sub

    Public Sub AddABall()
        If m_balls_live_target > 0 Then
            Log "Adding a ball to multiball: " & m_name
            m_balls_live_target = m_balls_live_target + 1
            m_balls_added_live = m_balls_added_live + 1
            m_queued_balls = m_queued_balls + 1
            glf_BIP = glf_BIP + 1
            SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me, m_queued_balls),Null), 1000
        End If
    End Sub

    Public Sub AddAballTimerStart()
        'Start the timer for add a ball ball save.
        'This is started when multiball add a ball is triggered if configured,
        'and the default timer is not still running.
        If m_shoot_again_enabled = True Then
            Exit Sub
        End If

        m_shoot_again_enabled = True

        Dim shoot_again_ms : shoot_again_ms = m_add_a_ball_shoot_again.Value()
        if shoot_again_ms = 0 Then
            'No shoot again. Just stop multiball right away
            StopMultiball()
            Exit Sub
        End If

        DispatchPinEvent "ball_save_" & m_configname & "_add_a_ball_timer_start", Null

        Dim grace_period_ms : grace_period_ms = m_add_a_ball_grace_period.Value()
        Dim hurry_up_time_ms : hurry_up_time_ms = m_add_a_ball_hurry_up_time.Value()
        StartShootAgain shoot_again_ms, grace_period_ms, hurry_up_time_ms
    End Sub

    Public Sub StopMultiball()
        '"""Stop shoot again."""
        Log("Stopping shoot again of multiball")
        m_shoot_again_enabled = False

        '# disable shoot again
        RemoveDelay m_name&"_disable_shoot_again"

        If m_grace_period_enabled Then
            RemoveDelay m_name&"_grace_period"
            RunGracePeriod()
        End If
        If m_hurry_up_enabled Then
            RemoveDelay m_name&"_hurry_up"
            RunHurryUp()
        End If
        Log "Stop Shoot Again, Queued Balls: " & QueuedBalls()
        'RemoveDelay m_name & "_queued_release"

        DispatchPinEvent m_name & "_shoot_again_ended", Null
    End Sub

    Public Function QueuedBalls()
        QueuedBalls = m_queued_balls
    End Function

    Public Function ReleaseQueuedBalls()
        m_queued_balls = m_queued_balls - 1
        Log "Queued Balls: " & m_queued_balls
        ReleaseQueuedBalls = m_queued_balls
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function MultiballsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set MultiballsHandler = kwargs
                Else
                    MultiballsHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "start"
            multiball.Start
        Case "reset"
            multiball.Reset
        Case "add_a_ball"
            multiball.AddABall
        Case "stop"
            multiball.StopMultiball
        Case "grace_period"
            multiball.RunGracePeriod
        Case "hurry_up"
            multiball.RunHurryUp
        Case "drain"
            kwargs = multiball.BallsDrained(kwargs)
        Case "queue_release"
            If multiball.QueuedBalls() > 0 Then
                If glf_plunger.HasBall = False And ballInReleasePostion = True And glf_plunger.IncomingBalls = 0 Then
                    Glf_ReleaseBall(Null)
                    SetDelay multiball.Name&"_auto_launch", "MultiballsHandler" , Array(Array("auto_launch", multiball),Null), 500
                    If multiball.ReleaseQueuedBalls() > 0 Then
                        SetDelay multiball.Name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", multiball), Null), 1000    
                    End If
                Else
                    SetDelay multiball.Name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", multiball), Null), 1000
                End If
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay multiball.Name&"_auto_launch", "MultiballsHandler" , Array(Array("auto_launch", multiball), Null), 500
            End If
    End Select

    If IsObject(args(1)) Then
        Set MultiballsHandler = kwargs
    Else
        MultiballsHandler = kwargs
    End If
End Function



Class GlfQueueEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "queue_event_player" : End Property

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "queue_event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Raw, newEvent
        'msgbox newEvent.Name
        m_eventValues.Add newEvent.Raw, value  
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_queue_event_player_play", "QueueEventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_queue_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        If m_events(evt).Evaluate() = False Then
            Exit Sub
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            Log "Dispatching Event: " & evtValue
            DispatchQueuePinEvent evtValue, Null
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_queue_event_player", message
        End If
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

Function QueueEventPlayerEventHandler(args)
    
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
        Set QueueEventPlayerEventHandler = kwargs
    Else
        QueueEventPlayerEventHandler = kwargs
    End If
End Function


Class GlfQueueRelayPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "queue_relay_player" : End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "queue_relay_player", Me)
        Set Init = Me
	End Function

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get EventName(name)
        If m_events.Exists(name) Then
            Set EventName = m_eventValues(name)
        Else
            Dim new_event : Set new_event = (new GlfEvent)(name)
            m_events.Add new_event.Raw, new_event
            Dim new_event_value : Set new_event_value = (new GlfQueueRelayEvent)()
            m_eventValues.Add new_event.Raw, new_event_value
            Set EventName = new_event_value
        End If
    End Property

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_queue_relay_player_play", "QueueRelayPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_queue_relay_player_play"
        Next
    End Sub

    Public Function FireEvent(evt)
        FireEvent=Empty
        If m_events(evt).Evaluate() Then
            'post a new event, and wait for the release
            DispatchPinEvent m_eventValues(evt).Post, Null
            FireEvent = m_eventValues(evt).WaitFor
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_queue_relay_player", message
        End If
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

Function QueueRelayPlayerEventHandler(args)
    
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
            Dim wait_for : wait_for = eventPlayer.FireEvent(ownProps(2))
            If Not IsEmpty(wait_for) Then
                kwargs.Add "wait_for", wait_for
            End If
    End Select
    If IsObject(args(1)) Then
        Set QueueRelayPlayerEventHandler = kwargs
    Else
        QueueRelayPlayerEventHandler = kwargs
    End If
End Function

Class GlfQueueRelayEvent

	Private m_wait_for, m_post
  
    Public Property Get WaitFor() : WaitFor = m_wait_for : End Property
    Public Property Let WaitFor(input) : m_wait_for = input : End Property
    Public Property Get Post() : Post = m_post : End Property
    Public Property Let Post(input) : m_post = input : End Property
        
	Public default Function init()
        m_wait_for = Empty
        m_post = Empty
	    Set Init = Me
	End Function

End Class


Class GlfRandomEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "random_event_player" : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get EventNames() : EventNames = m_events.Keys : End Property
    Public Property Get EventName(value)
        
        Dim newEvent : Set newEvent = (new GlfEvent)(value)
        m_events.Add newEvent.Raw, newEvent
        Dim newRandomEvent : Set newRandomEvent = (new GlfRandomEvent)(value, m_mode, UBound(m_events.Keys))
        m_eventValues.Add newEvent.Raw, newRandomEvent
        
        Set EventName = newRandomEvent
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "random_event_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        Log "Activating"
        For Each evt In m_events.Keys()
            Log "Adding: " & m_events(evt).EventName & ". For Key: " & m_mode & "_" & evt & "_random_event_player_play"
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_random_event_player_play", "RandomEventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_random_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        Log "Firing Random Event:  " & evt
        If m_events(evt).Evaluate() Then
            Dim event_to_fire
            event_to_fire = m_eventValues(evt).GetNextRandomEvent()
            If Not IsEmpty(event_to_fire) Then
                Log "Dispatching Event: " & event_to_fire
                DispatchPinEvent event_to_fire, Null
            Else
                Log "No event available to fire"
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_random_event_player", message
        End If
    End Sub

    Public Function ToYaml()
            Dim yaml
        Dim evt
        If UBound(m_eventValues.Keys) > -1 Then
            For Each key in m_eventValues.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_eventValues(key).ToYaml 
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function RandomEventPlayerEventHandler(args)
    
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
        Set RandomEventPlayerEventHandler = kwargs
    Else
        RandomEventPlayerEventHandler = kwargs
    End If
End Function

Class GlfSegmentDisplayPlayer

    Private m_priority
    Private m_mode
    Private m_name
    Private m_debug
    Private m_events
    private m_base_device

    Public Property Get Name() : Name = "segment_player" : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    
    Public Property Get EventName(value)
        Dim newEvent : Set newEvent = (new GlfEvent)(value)

        If m_events.Exists(newEvent.Raw) Then
            Set EventName = m_events(newEvent.Raw)
        Else
            Dim new_segment_event : Set new_segment_event = (new GlfSegmentDisplayPlayerEvent)(newEvent)
            m_events.Add newEvent.Raw, new_segment_event
            Set EventName = new_segment_event
        End If
    End Property

	Public default Function init(mode)
        m_name = "segment_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "segment_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).GlfEvent.EventName, m_mode & "_" & evt & "_segment_player_play", "SegmentPlayerEventHandler", m_priority+m_events(evt).GlfEvent.Priority, Array("play", Me, m_events(evt), m_events(evt).GlfEvent.EventName)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt, display
        Dim displays_to_update : Set displays_to_update = CreateObject("Scripting.Dictionary")
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).GlfEvent.EventName, m_mode & "_" & evt & "_segment_player_play"
            Set displays_to_update = PlayOff(m_events(evt).GlfEvent.EventName, m_events(evt), displays_to_update)
        Next
        
        For Each display in displays_to_update.Keys()
            glf_segment_displays(display).UpdateStack()
        Next
    End Sub

    Public Sub Play(evt, segment_event)
        Dim i
        Dim segment_event_displays : segment_event_displays = segment_event.Displays()
        For i=0 to UBound(segment_event_displays)
            SegmentPlayerCallbackHandler evt, segment_event_displays(i), m_mode, m_priority
        Next
    End Sub

    Public Function PlayOff(evt, segment_event, displays_to_update)
        Dim i, segment_item
        Dim segment_event_displays : segment_event_displays = segment_event.Displays()
        For i=0 to UBound(segment_event_displays)
            Set segment_item = segment_event_displays(i)
            Dim key
            key = m_mode & "." & "segment_player_player." & segment_item.Display
            If Not IsEmpty(segment_item.Key) Then
                key = key & segment_item.Key
            End If
            Dim display : Set display = glf_segment_displays(segment_item.Display)
            RemoveDelay key
            display.RemoveTextByKeyNoUpdate key
            If Not displays_to_update.Exists(segment_item.Display) Then
                displays_to_update.Add segment_item.Display, True
            End If 
        Next
        Set PlayOff = displays_to_update
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
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

Class GlfSegmentDisplayPlayerEvent

    Private m_items
    Private m_event

    Public Property Get Displays() : Displays = m_items.Items() : End Property

    Public Property Get GlfEvent() : Set GlfEvent = m_event : End Property

    Public Property Get Display(value)
        If m_items.Exists(value) Then
            Set Display = m_items(value)
        Else
            Dim new_item : Set new_item = (new GlfSegmentPlayerEventItem)()
            new_item.Display = value
            m_items.add value, new_item
            Set Display = new_item
        End If
    End Property

    Public default Function init(evt)
        Set m_items = CreateObject("Scripting.Dictionary")
        Set m_event = evt
        Set Init = Me
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
            Set Text = m_text
        Else
            Text = Null
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
    Public Property Let Expire(input) : m_expire = input : End Property

    Public Property Get FlashMask() : FlashMask = m_flash_mask : End Property
    Public Property Let FlashMask(input) : m_flash_mask = input : End Property
                        
    Public Property Get Flashing() : Flashing = m_flashing : End Property
    Public Property Let Flashing(input) : m_flashing = input : End Property
                            
    Public Property Get Key() : Key = m_key : End Property
    Public Property Let Key(input) : m_key = input : End Property

    Public Property Get Color() : Color = m_color : End Property
    Public Property Let Color(input) : m_color = input : End Property

    Public Property Get HasTransition() : HasTransition = Not IsNull(m_transition) : End Property    
    Public Property Get HasTransitionOut() : HasTransitionOut = Not IsNull(m_transition_out) : End Property

    Public Property Get Transition()
        If IsNull(m_transition) Then
            Set m_transition = (new GlfSegmentPlayerTransition)("bee")
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
        m_flashing = "no_flash"
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
    
    Public Property Get Text()
        Text = m_text
    End Property
    Public Property Let Text(input)
        m_text = input
    End Property

    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let Direction(input) : m_direction = input : End Property                          

	Public default Function Init(loo)
        m_type = "push"
        m_text = Empty
        m_direction = "right"
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
            If ownProps(2).GlfEvent.Evaluate() Then
                SegmentPlayer.Play ownProps(3), ownProps(2)
            End If
        Case "remove"
            RemoveDelay ownProps(2)
            SegmentPlayer.RemoveTextByKey ownProps(2)
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
            Dim transition, transition_out : transition = Null : transition_out = Null
            If segment_item.HasTransition() Then
                Set transition = segment_item.Transition
            End If
            If segment_item.HasTransitionOut() Then
                Set transition_out = segment_item.TransitionOut
            End If
            display.AddTextEntry segment_item.Text, segment_item.Color, segment_item.Flashing, segment_item.FlashMask, transition, transition_out, priority + segment_item.Priority, key
                                
            If segment_item.Expire > 0 Then
                SetDelay key & "_expire", "SegmentPlayerEventHandler",  Array(Array("remove", display, key)), segment_item.Expire
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

Class GlfSequenceShots

    Private m_name
    Private m_command_name
    Private m_lock_device
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_cancel_events
    Private m_cancel_switches
    Private m_delay_event_list
    Private m_delay_switch_list
    Private m_event_sequence
    Private m_sequence_timeout
    Private m_switch_sequence
    Private m_start_event
    Private m_sequence_count
    Private m_active_delays
    Private m_active_sequences
    Private m_sequence_events
    Private m_start_time

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

    Public Property Let CancelEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_cancel_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let CancelSwitches(value): m_cancel_switches = value: End Property
    Public Property Let DelayEventList(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_delay_event_list.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let DelaySwitchList(value): m_delay_switch_list = value: End Property
    Public Property Let EventSequence(value)
        m_event_sequence = value
        If m_sequence_count = 0 Then
            m_sequence_events = value
        Else
            Redim Preserve m_sequence_events(m_sequence_count+UBound(value))
            Dim i
            For i = 0 To UBound(value)
                m_sequence_events(m_sequence_count + i) = value(i)
            Next
        End If
        m_start_event = value(0)
        m_sequence_count = m_sequence_count + (UBound(m_sequence_events)+1)
    End Property
    Public Property Let SequenceTimeout(value): Set m_sequence_timeout = CreateGlfInput(value): End Property
    Public Property Let SwitchSequence(value)
        m_switch_sequence = value
        If m_sequence_count = 0 Then
            m_start_event = value(0) & "_active"
        End If
        Redim Preserve m_sequence_events(m_sequence_count+UBound(value))
        Dim i
        For i = 0 To UBound(value)
            m_sequence_events(m_sequence_count + i) = value(i) & "_active"
        Next
        m_sequence_count = m_sequence_count + (UBound(m_sequence_events)+1)
    End Property

	Public default Function init(name, mode)
        m_name = "sequence_shot_" & name
        m_command_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        Set m_cancel_events = CreateObject("Scripting.Dictionary")
        Set m_delay_event_list = CreateObject("Scripting.Dictionary")
        Set m_active_sequences = CreateObject("Scripting.Dictionary")
        Set m_active_delays = CreateObject("Scripting.Dictionary")
        
        m_sequence_events = Array()
        m_cancel_switches = Array()
        m_start_time = 0
        m_event_sequence = Array()
        m_switch_sequence = Array()
        Set m_sequence_timeout = CreateGlfInput(0)
        m_sequence_count = 0
        m_start_event = Empty
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "sequence_shot_", Me)
        
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
        For Each evt in m_event_sequence
            AddPinEventListener evt, m_name & "_" & evt & "_advance", "SequenceShotsHandler", m_priority, Array("advance", Me, evt)
        Next
        For Each evt in m_switch_sequence
            AddPinEventListener evt & "_active", m_name & "_" & evt & "_advance", "SequenceShotsHandler", m_priority, Array("advance", Me, evt & "_active")
        Next
        For Each evt in m_cancel_events.Keys
            AddPinEventListener m_cancel_events(evt).EventName, m_name & "_" & evt & "_cancel", "SequenceShotsHandler", m_priority+m_cancel_events(evt).Priority, Array("cancel_event", Me, m_cancel_events(evt))
        Next
        For Each evt in m_delay_event_list.Keys
            AddPinEventListener m_delay_event_list(evt).EventName, m_name & "_" & evt & "_delay", "SequenceShotsHandler", m_priority+m_delay_event_list(evt).Priority, Array("delay_event", Me, m_delay_event_list(evt))
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_event_sequence
            RemovePinEventListener evt, m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_switch_sequence
            RemovePinEventListener evt & "_active", m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_cancel_events.Keys
            RemovePinEventListener m_cancel_events(evt).EventName, m_name & "_" & evt & "_cancel"
        Next
        For Each evt in  m_delay_event_list.Keys
            RemovePinEventListener m_delay_event_list(evt).EventName, m_name & "_" & evt & "_delay"
        Next
    End Sub

    Sub SequenceAdvance(event_name)
        ' Since we can track multiple simultaneous sequences (e.g. two balls
        ' going into an orbit in a row), we first have to see whether this
        ' switch is starting a new sequence or continuing an existing one

        Log "Sequence advance: " & event_name

        If event_name = m_start_event Then
            If m_sequence_count > 1 Then
                ' start a new sequence
                StartNewSequence()
            ElseIf UBound(m_active_delays.Keys) = -1 Then
                ' if it only has one step it will finish right away
                Completed()
            End If
        Else
            ' Get the seq_id of the first sequence this switch is next for.
            ' This is not a loop because we only want to advance 1 sequence
            Dim k, seq
            seq = Null
            For Each k In m_active_sequences.Keys
                Log m_active_sequences(k).NextEvent
                If m_active_sequences(k).NextEvent = event_name Then
                    Set seq = m_active_sequences(k)
                    Exit For
                End If
            Next

            If Not IsNull(seq) Then
                ' advance this sequence
                AdvanceSequence(seq)
            End If
        End If
    End Sub

    Public Sub StartNewSequence()
        ' If the sequence hasn't started, make sure we're not within the
        ' delay_switch hit window

        If UBound(m_active_delays.Keys)>-1 Then
            Log "There's a delay timer in effect. Sequence will not be started."
            Exit Sub
        End If

        'record start time
        m_start_time = gametime

        ' create a new sequence
        Dim seq_id : seq_id = "seq_" & glf_SeqCount
        glf_SeqCount = glf_SeqCount + 1

        Dim next_event : next_event = m_sequence_events(1)

        Log "Setting up a new sequence. Next: " & next_event

        m_active_sequences.Add seq_id, (new GlfActiveSequence)(seq_id, 0, next_event)

        ' if this sequence has a time limit, set that up
        If m_sequence_timeout.Value > 0 Then
            Log "Setting up a sequence timer for " & m_sequence_timeout.Value
            SetDelay seq_id, "SequenceShotsHandler" , Array(Array("seq_timeout", Me, seq_id),Null), m_sequence_timeout.Value
        End If
    End Sub

    Public Sub AdvanceSequence(sequence)
        ' Remove this sequence from the list
        If sequence.CurrentPositionIndex = (m_sequence_count - 2) Then  ' complete
            Log "Sequence complete!"
            RemoveDelay sequence.SeqId
            m_active_sequences.Remove sequence.SeqId
            Completed()
        Else
            Dim current_position_index : current_position_index = sequence.CurrentPositionIndex + 1
            Dim next_event : next_event = m_sequence_events(current_position_index + 1)
            Log "Advancing the sequence. Next: " & next_event
            sequence.CurrentPositionIndex = current_position_index
            sequence.NextEvent = next_event
        End If
    End Sub

    Public Sub Completed()
        'measure the elapsed time between start and completion of the sequence
        Dim elapsed
        If m_start_time > 0 Then
            elapsed = gametime - m_start_time
        Else
            elapsed = 0
        End If

        'Post sequence complete event including its elapsed time to complete.
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "elapsed", elapsed
        End With
        DispatchPinEvent m_command_name & "_hit", kwargs
    End Sub

    Public Sub ResetAllSequences()
        'Reset all sequences."""
        Dim k
        For Each k in m_active_sequences.Keys
            RemoveDelay m_active_sequences(k).SeqId
        Next

        m_active_sequences.RemoveAll()
    End Sub

    Public Sub DelayEvent(delay, name)
        Log "Delaying sequence by " & delay
        SetDelay name & "_delay_timer", "SequenceShotsHandler" , Array(Array("release_delay", Me, name),Null), delay
        m_active_delays.Add name, True
    End Sub

    Public Sub ReleaseDelay(name)
        m_active_delays.Remove name
    End Sub

    Public Sub FireSequenceTimeout(seq_id)
        Log "Sequence " & seq_id & " timeouted"
        m_active_sequences.Remove seq_id
        DispatchPinEvent m_name & "_timeout", Null
    End Sub

     Public Function ToYaml()
        Dim yaml, key, x
        yaml = "  " & Replace(m_name, "sequence_shot_", "") & ":" & vbCrLf
        If UBound(m_switch_sequence) > -1 Then
            yaml = yaml & "    switch_sequence: " & vbCrLf
            For Each key in m_switch_sequence
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_event_sequence) > -1 Then
            yaml = yaml & "    event_sequence: " & vbCrLf
            For Each key in m_event_sequence
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_cancel_events.Keys) > -1 Then
            yaml = yaml & "    cancel_events: " & vbCrLf
            For Each key in m_cancel_events.Keys
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_cancel_switches) > -1 Then
            yaml = yaml & "    cancel_swithces: " & vbCrLf
            For Each key in m_cancel_switches
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_delay_event_list.Keys()) > -1 Then
            yaml = yaml & "    delay_event_list: " & vbCrLf
            For Each key in m_delay_event_list.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If m_sequence_timeout.Raw <> 0 Then
            yaml = yaml & "    sequence_timeout: " & m_sequence_timeout.Raw & vbCrLf
        End If

        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Class GlfActiveSequence

    Private m_next_event, m_seq_id, m_idx

    Public Property Get SeqId(): SeqId = m_seq_id: End Property
    Public Property Get NextEvent(): NextEvent = m_next_event: End Property
    Public Property Let NextEvent(value): m_next_event = value: End Property
    Public Property Get CurrentPositionIndex(): CurrentPositionIndex = m_idx: End Property
    Public Property Let CurrentPositionIndex(value): m_idx = value: End Property

    Public default Function init(seq_id, idx, next_event)
        m_seq_id = seq_id
        m_idx = idx
        m_next_event = next_event
        Set Init = Me
    End Function

End Class

Function SequenceShotsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim sequence_shot : Set sequence_shot = ownProps(1)
    Select Case evt
        Case "advance"
            sequence_shot.SequenceAdvance ownProps(2)
        Case "cancel"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            sequence_shot.ResetAllSequences
        Case "delay"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            sequence_shot.DelayEvent glfEvent.Delay, glfEvent.EventName
        Case "seq_timeout"
            sequence_shot.FireSequenceTimeout ownProps(2)
        Case "release_delay"
            sequence_shot.ReleaseDelay ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set SequenceShotsHandler = kwargs
    Else
        SequenceShotsHandler = kwargs
    End If
End Function
Class GlfShotGroup
    Private m_name
    Private m_mode
    Private m_priority
    private m_base_device
    private m_debug
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
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

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
            m_enable_rotation_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let DisableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_rotation_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RotateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RotateLeftEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_left_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RotateRightEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_right_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
 
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
            AddPinEventListener m_enable_rotation_events(evt).EventName, m_name & "_" & evt & "_enable_rotation", "ShotGroupEventHandler", m_priority+m_enable_rotation_events(evt).Priority, Array("enable_rotation", Me, m_enable_rotation_events(evt))
        Next
        For Each evt in m_disable_rotation_events.Keys()
            AddPinEventListener m_disable_rotation_events(evt).EventName, m_name & "_" & evt & "_disable_rotation", "ShotGroupEventHandler", m_priority+m_disable_rotation_events(evt).Priority, Array("disable_rotation", Me, m_disable_rotation_events(evt))
        Next
        For Each evt in m_restart_events.Keys()
            AddPinEventListener m_restart_events(evt).EventName, m_name & "_" & evt & "_restart", "ShotGroupEventHandler", m_priority+m_restart_events(evt).Priority, Array("restart", Me, m_restart_events(evt))
        Next
        For Each evt in m_reset_events.Keys()
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset", "ShotGroupEventHandler", m_priority+m_reset_events(evt).Priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_rotate_events.Keys()
            AddPinEventListener m_rotate_events(evt).EventName, m_name & "_" & evt & "_rotate", "ShotGroupEventHandler", m_priority+m_rotate_events(evt).Priority, Array("rotate", Me, m_rotate_events(evt))
        Next
        For Each evt in m_rotate_left_events.Keys()
            AddPinEventListener m_rotate_left_events(evt).EventName, m_name & "_" & evt & "_rotate_left", "ShotGroupEventHandler", m_priority+m_rotate_left_events(evt).Priority, Array("rotate_left", Me, m_rotate_left_events(evt))
        Next
        For Each evt in m_rotate_right_events.Keys()
            AddPinEventListener m_rotate_right_events(evt).EventName, m_name & "_" & evt & "_rotate_right", "ShotGroupEventHandler", m_priority+m_rotate_right_events(evt).Priority, Array("rotate_right", Me, m_rotate_right_events(evt))
        Next
        Dim shot_name
        For Each shot_name in m_shots
            AddPinEventListener shot_name & "_hit", m_name & "_" & m_mode & "_hit", "ShotGroupEventHandler", m_priority, Array("hit", Me, Null)
            AddPlayerStateEventListener "shot_" & shot_name, m_name & "_" & m_mode & "_complete", -1, "ShotGroupEventHandler", m_priority, Array("complete", Me, Null)
        Next
    End Sub
 
    Public Sub Deactivate
        Dim evt
        m_rotation_enabled = True
        For Each evt in m_enable_rotation_events.Keys()
            RemovePinEventListener m_enable_rotation_events(evt).EventName, m_name & "_" & evt & "_enable_rotation"
        Next
        For Each evt in m_disable_rotation_events.Keys()
            RemovePinEventListener m_disable_rotation_events(evt).EventName, m_name & "_" & evt & "_disable_rotation"
        Next
        For Each evt in m_restart_events.Keys()
            RemovePinEventListener m_restart_events(evt).EventName, m_name & "_" & evt & "_restart"
        Next
        For Each evt in m_reset_events.Keys()
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_rotate_events.Keys()
            RemovePinEventListener m_rotate_events(evt).EventName, m_name & "_" & evt & "_rotate"
        Next
        For Each evt in m_rotate_left_events.Keys()
            RemovePinEventListener m_rotate_left_events(evt).EventName, m_name & "_" & evt & "_rotate_left"
        Next
        For Each evt in m_rotate_right_events.Keys()
            RemovePinEventListener m_rotate_right_events(evt).EventName, m_name & "_" & evt & "_rotate_right"
        Next
        Dim shot_name
        For Each shot_name in m_shots
            RemovePinEventListener shot_name & "_hit", m_name & "_" & m_mode & "_hit"
            RemovePlayerStateEventListener "shot_" & shot_name, m_name & "_" & m_mode & "_complete"
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
        Log "Enabling"
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
        Log "Enabling Rotation"
        m_rotation_enabled = True
    End Sub
 
    Public Sub DisableRotation
        Log "Disabling Rotation"
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
            Log "Rotating Shot:" & shot
            m_temp_shots(shot).Jump shot_states(x), True, False
            x=x+1
        Next 
        m_isRotating = False
        CheckForComplete()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
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
    If Not IsNull(ownProps(2)) Then
        Dim glf_event : Set glf_event = ownProps(2)
        If glf_event.Evaluate() = False Then
            If IsObject(args(1)) Then
                Set ShotGroupEventHandler = kwargs
            Else
                ShotGroupEventHandler = kwargs
            End If
            Exit Function
        End If
    End If
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
            If UBound(m_states(evt).EventsWhenCompleted) > -1 Then
                yaml = yaml & "       events_when_completed: " & Join(m_states(evt).EventsWhenCompleted, ",") & vbCrLf
            End If

            If Ubound(state.Tokens().Keys)>-1 Then
                yaml = yaml & "       show_tokens: " & vbCrLf
                Dim state_tokens : Set state_tokens = state.Tokens()
                For Each token in state_tokens.Keys()
                    yaml = yaml & "         " & token & ": " & state_tokens(token) & vbCrLf
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
    Private m_debug
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
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
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
    Public Property Get ControlEvents()
            Dim control_event_count : control_event_count = UBound(m_control_events.Keys)
            Dim newEvent : Set newEvent = (new GlfShotControlEvent)()
            m_control_events.Add CStr(control_event_count+1), newEvent
            Set ControlEvents = newEvent
    End Property
    Public Property Let AdvanceEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_advance_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Raw, newEvent
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
            m_hit_events.Add newEvent.Raw, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_internal_cache_id = -1
        m_enabled = False
        m_persist = True
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot", Me)
        m_debug = False
        m_profile = "default"
        m_player_var_name = "shot_" & m_name
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
        If GetPlayerState(m_player_var_name) = False Then
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
            RemovePinEventListener m_hit_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_hit"
        Next
        For Each evt in m_advance_events.Keys
            RemovePinEventListener m_advance_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events().Keys
                Dim control_events_events : Set control_events_events = m_control_events(evt).Events()
                RemovePinEventListener control_events_events(cEvt).EventName, m_mode & "_" & m_name & "_control_" & cEvt
            Next
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener m_reset_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_restart_events.Keys
            RemovePinEventListener m_restart_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_restart"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_switches
            AddPinEventListener evt & "_active", m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me, Null)
        Next
        For Each evt in m_hit_events.Keys
            AddPinEventListener m_hit_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_hit", "ShotEventHandler", m_priority, Array("hit", Me, m_hit_events(evt))
        Next
        For Each evt in m_advance_events.Keys
            AddPinEventListener m_advance_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_advance", "ShotEventHandler", m_priority, Array("advance", Me, m_advance_events(evt))
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events().Keys
                Dim control_events_events : Set control_events_events = m_control_events(evt).Events()
                AddPinEventListener control_events_events(cEvt).EventName, m_mode & "_" & m_name & "_control_" & cEvt, "ShotEventHandler", m_priority+control_events_events(cEvt).Priority, Array("control", Me, control_events_events(cEvt), m_control_events(evt))
            Next
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener m_reset_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_reset", "ShotEventHandler", m_priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_restart_events.Keys
            AddPinEventListener m_restart_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_restart", "ShotEventHandler", m_priority, Array("restart", Me, m_restart_events(evt))
        Next
        'Play the show for the active state
        PlayShowForState(m_state)
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_hit_events.Keys
            RemovePinEventListener m_hit_events(evt).EventName, m_mode & "_" & m_name & "_" & evt  & "_hit"
        Next
        Dim x
        For x=0 to Glf_ShotProfiles(m_profile).StatesCount()
            StopShowForState(x)
        Next
    End Sub

    Private Sub StopShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        Log "Removing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        If glf_running_shows.Exists(m_mode & "_" & CStr(state) & "_" & m_name & "_" & profileState.Key) Then 
            glf_running_shows(m_mode & "_" & CStr(state) & "_" & m_name & "_" & profileState.Key).StopRunningShow()
        End If
    End Sub

    Private Sub PlayShowForState(state)
        If m_enabled = False Then
            Exit Sub
        End If
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        Log "Playing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
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
        Log "Hit! Profile: "&m_profile&", State: " & profile.StateName(m_state)

        Dim advancing
        If profile.AdvanceOnHit Then
            Log "Advancing shot because advance_on_hit is True."
            advancing = Advance(False)
        Else
            Log "Not advancing shot"
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

        Log "Advancing 1 step. Profile: "&m_profile&", Current State: " & profile.StateName(m_state)

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
        Log "Received jump request. State: " & state & ", Force: "& force

        If Not m_enabled And Not force Then
            Log "Profile is disabled and force is False. Not jumping"
            Exit Sub
        End If
        If state = m_state And Not force_show Then
            Log "Shot is already in the jump destination state"
            Exit Sub
        End If
        Log "Jumping to profile state " & state

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

    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
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

        Dim enable_events_keys : enable_events_keys = m_base_device.EnableEvents().Keys
        Dim enable_events : Set enable_events = m_base_device.EnableEvents()
        If UBound(enable_events_keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in enable_events_keys
                yaml = yaml & enable_events(key).Raw
                If x <> UBound(enable_events_keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        Dim disable_events_keys : disable_events_keys = m_base_device.DisableEvents().Keys
        Dim disable_events : Set disable_events = m_base_device.DisableEvents()
        If UBound(disable_events_keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in disable_events_keys
                yaml = yaml & disable_events(key).Raw
                If x <> UBound(disable_events_keys) Then
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
    If Not IsNull(ownProps(2)) Then
        Dim glf_event : Set glf_event = ownProps(2)
        If glf_event.Evaluate() = False Then
            If IsObject(args(1)) Then
                Set ShotEventHandler = kwargs
            Else
                ShotEventHandler = kwargs
            End If
            Exit Function
        End If
    End If
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
            shot.Jump ownProps(3).State, ownProps(3).Force, ownProps(3).ForceShow
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
  
	Public Property Get Events(): Set Events = m_events: End Property
    Public Property Let Events(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Get State(): State = m_state End Property
    Public Property Let State(input): m_state = input End Property

    Public Property Get Force(): Force = m_force: End Property
	Public Property Let Force(input): m_force = input: End Property
  
	Public Property Get ForceShow(): ForceShow = m_force_show: End Property
	Public Property Let ForceShow(input): m_force_show = input: End Property   

	Public default Function init()
        Set m_events = CreateObject("Scripting.Dictionary")
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
    Private m_eventValues
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "show_player" : End Property
    Public Property Get EventShows() : EventShows = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_show : Set new_show = (new GlfShowPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_show
        Set EventName = new_show
        
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "show_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_show_player_play", "ShowPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_show_player_play"
            PlayOff m_eventValues(evt).Key
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            If m_eventValues(evt).Action = "stop" Then
                PlayOff m_eventValues(evt).Key
            Else
                Dim new_running_show
                Set new_running_show = (new GlfRunningShow)(m_name & "_" & m_eventValues(evt).Key, m_eventValues(evt).Key, m_eventValues(evt), m_priority, Null, Null)
                If m_eventValues(evt).BlockQueue = True Then
                    Play = m_name & "_" & m_eventValues(evt).Key & "_" & m_eventValues(evt).Key  & "_unblock_queue"
                End If
            End If
        End If
    End Function

    Public Sub PlayOff(key)
        If glf_running_shows.Exists(m_name & "_" & key) Then 
            glf_running_shows(m_name & "_" & key).StopRunningShow()
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_eventValues.Keys) > -1 Then
            For Each key in m_eventValues.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_eventValues(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function ShowPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim ShowPlayer : Set ShowPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            ShowPlayer.Activate
        Case "deactivate"
            ShowPlayer.Deactivate
        Case "play"
            Dim block_queue
            block_queue = ShowPlayer.Play(ownProps(2))
            If Not IsEmpty(block_queue) Then
                kwargs.Add "wait_for", block_queue
            End If
    End Select
    If IsObject(args(1)) Then
        Set ShowPlayerEventHandler = kwargs
    Else
        ShowPlayerEventHandler = kwargs
    End If
End Function

Class GlfShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms, m_duration, m_priority, m_internal_cache_id
    Private m_block_queue
    Private m_events_when_completed
    Private m_color_lookup
  
	Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property

    Public Property Get ColorLookup(): ColorLookup = m_color_lookup End Property
    Public Property Let ColorLookup(input): m_color_lookup = input End Property

    Public Property Get Show()
        If IsNull(m_show) Then
            Show = Null
        Else
            Set Show = m_show
        End If
    End Property
	Public Property Let Show(input)
        'msgbox "input:" & input
        If glf_shows.Exists(input) Then
            Set m_show = glf_shows(input)
        End If
    End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property      
        
    Public Property Get BlockQueue(): BlockQueue = m_block_queue : End Property
    Public Property Let BlockQueue(input): m_block_queue = input : End Property
    
    Public Property Get EventsWhenCompleted(): EventsWhenCompleted = m_events_when_completed : End Property
    Public Property Let EventsWhenCompleted(input): m_events_when_completed = input: End Property
            
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

        If UBound(m_events_when_completed) > -1 Then
            yaml = yaml & "      events_when_completed: " & Join(m_events_when_completed, ",") & vbCrLf
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
        m_color_lookup = Empty
        m_block_queue = False
        m_events_when_completed = Array()
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

    Public Function StepAtIndex(index)
        Dim step_at_index_items : step_at_index_items = m_steps.Items()
        Set StepAtIndex = step_at_index_items(index)
    End Function
    
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
            Dim steps_items : steps_items = m_steps.Items()
            Dim prevStep : Set prevStep = steps_items(UBound(m_steps.Keys()))
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
        Dim yaml, show_step
        For Each show_step in m_steps.Items()
            If Not IsNull(show_step.Duration) Then
                yaml = yaml & "- duration: " & show_step.Duration & "s" & vbCrLf
            Else
                If Not IsNull(show_step.AbsoluteTime) Then
                    yaml = yaml & "- time: " & show_step.AbsoluteTime & "s" & vbCrLf
                End If
                If Not IsNull(show_step.RelativeTime) Then
                    yaml = yaml & "- time: +" & show_step.RelativeTime & "s" & vbCrLf
                End If
            End If
            
            yaml = yaml & show_step.toYaml() & vbCrLf
        Next
        ToYaml = yaml
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
    Private m_loops
    Private m_shows_added

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

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property
        
    Public Property Get ShowSettings(): Set ShowSettings = m_show_settings: End Property
    Public Property Let ShowSettings(input)
        Set m_show_settings = input
        m_loops = m_show_settings.Loops
    End Property

    Public Property Get ShowsAdded()
        If IsNull(m_shows_added) Then
            ShowsAdded = Null
        Else
            Set ShowsAdded = m_shows_added
        End If
    End Property

    Public Sub SetShowsAdded(shows)
        Set m_shows_added = shows
    End Sub

    Public Sub ClearShowsAdded()
        m_shows_added = Null
    End Sub
    
    Public default Function init(rname, rkey, show_settings, priority, tokens, cache_id)
        m_show_name = rname
        m_key = rkey
        m_current_step = 0
        m_priority = priority
        m_internal_cache_id = cache_id
        m_loops=show_settings.Loops
        Set m_show_settings = show_settings
        m_shows_added = Null
        Dim key
        Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
        If Not IsNull(m_show_settings.Tokens) Then
            Dim show_settings_tokens : Set show_settings_tokens = m_show_settings.Tokens()
            For Each key In show_settings_tokens.Keys()
                mergedTokens.Add key, show_settings_tokens(key)
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
            msgbox "show " & running_show.CacheName & " not cached! Problem with caching"
        End If
        Dim lightStack
        For Each light in cached_show_lights.Keys()

            Set lightStack = glf_lightStacks(light)
            
            If Not lightStack.IsEmpty() Then
                lightStack.PopByKey(m_show_name & "_" & m_key)
            End If

            Dim show_key
            For Each show_key in glf_running_shows.Keys
                If Left(show_key, Len("fade_" & m_show_name & "_" & m_key & "_" & light)) = "fade_" & m_show_name & "_" & m_key & "_" & light Then
                    glf_running_shows(show_key).StopRunningShow()
                End If
            Next
            
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
        If glf_debug_level = "Debug" Then
            glf_debugLog.WriteToLog "Running Show", message
        End If
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
            msgbox running_show.CacheName & " show not cached! Problem with caching"
        End If

        If Not IsNull(running_show.ShowsAdded()) Then
            Dim show_added
            For Each show_added in running_show.ShowsAdded().Keys()
                If glf_running_shows.Exists(show_added) Then 
                    glf_running_shows(show_added).StopRunningShow()
                End If
            Next
            running_show.ClearShowsAdded()
        End If  

        Dim shows_added, replacement_color
        replacement_color = Empty
        If Not IsEmpty(running_show.ShowSettings.ColorLookup) Then
            Dim show_settings_color_lookup : show_settings_color_lookup = running_show.ShowSettings.ColorLookup()
            replacement_color = show_settings_color_lookup(running_show.CurrentStep)
        End If
        shows_added = LightPlayerCallbackHandler(running_show.Key, Array(cached_show_seq(running_show.CurrentStep)), running_show.ShowName, running_show.Priority + running_show.ShowSettings.Priority, True, running_show.ShowSettings.Speed, replacement_color)
        If Not IsNull(shows_added(0)) Then
            'Fade shows were added, log them agains the current show.
            running_show.SetShowsAdded(shows_added(0))
        End If
    End If
    If UBound(nextStep.ShowsInStep().Keys())>-1 Then
        Dim show_item
        Dim show_items : show_items = nextStep.ShowsInStep().Items()
        For Each show_item in show_items
            If show_item.Action = "stop" Then
                If glf_running_shows.Exists(running_show.Key & "_" & show_item.Show & "_" & show_item.Key) Then 
                    glf_running_shows(running_show.Key & "_" & show_item.Show & "_" & show_item.Key).StopRunningShow()
                End If
            Else
                Dim new_running_show
                'MsgBox running_show.Priority + running_show.ShowSettings.Priority
                'msgbox running_show.Key & "_" & show_item.Key
                Set new_running_show = (new GlfRunningShow)(show_item.Key, show_item.Key, show_item, running_show.Priority + running_show.ShowSettings.Priority, Null, Null)
            End If
        Next
    End If
    If UBound(nextStep.DOFEventsInStep().Keys())>-1 Then
        Dim dof_item
        Dim dof_items : dof_items = nextStep.DOFEventsInStep().Items()
        For Each dof_item in dof_items
            DOF dof_item.DOFEvent, dof_item.Action
        Next
    End If
    If UBound(nextStep.SlidesInStep().Keys())>-1 Then
        Dim slide_item
        Dim slide_items : slide_items = nextStep.SlidesInStep().Items()
        For Each slide_item in slide_items
            
        Next
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
        If running_show.Loops = -1 Or running_show.Loops > 1 Then
            If running_show.Loops > 1 Then
                running_show.Loops = running_show.Loops - 1
            End If
            running_show.CurrentStep = 0
            SetDelay running_show.ShowName & "_" & running_show.Key, "GlfShowStepHandler", Array(running_show), (nextStep.Duration / running_show.ShowSettings.Speed) * 1000
        Else
'            glf_debugLog.WriteToLog "Running Show", "STOPPING SHOW, NO Loops"
            If UBound(running_show.ShowSettings().EventsWhenCompleted) > -1 Then
                Dim evt_when_completed
                For Each evt_when_completed in running_show.ShowSettings().EventsWhenCompleted
                    DispatchPinEvent evt_when_completed, Null
                Next
            End If
            DispatchPinEvent running_show.ShowName & "_" & running_show.Key & "_unblock_queue", Null
            running_show.StopRunningShow()
        End If
    Else
'        glf_debugLog.WriteToLog "Running Show", "Scheduling Next Step"
        SetDelay running_show.ShowName & "_" & running_show.Key, "GlfShowStepHandler", Array(running_show), (nextStep.Duration / running_show.ShowSettings.Speed) * 1000
    End If
End Function

Class GlfShowStep

    Private m_lights, m_shows, m_dofs, m_slides, m_time, m_duration, m_isLastStep, m_absTime, m_relTime

    Public Property Get Lights(): Lights = m_lights: End Property
    Public Property Let Lights(input) : m_lights = input: End Property

    Public Property Get ShowsInStep(): Set ShowsInStep = m_shows: End Property
    Public Property Get Shows(name)
        Dim new_show : Set new_show = (new GlfShowPlayerItem)()
        new_show.Show = name
        m_shows.Add name & CStr(UBound(m_shows.Keys)), new_show
        Set Shows = new_show
    End Property

    Public Property Get DOFEventsInStep(): Set DOFEventsInStep = m_dofs: End Property
    Public Property Get DOFEvent(dof_event)
        Dim new_dof : Set new_dof = (new GlfDofPlayerItem)()
        new_dof.DOFEvent = dof_event
        m_dofs.Add dof_event & CStr(UBound(m_dofs.Keys)), new_dof
        Set DOFEvent = new_dof
    End Property

    Public Property Get SlidesInStep(): Set SlidesInStep = m_slides: End Property
        Public Property Get Slides(slide)
            Dim new_slide : Set new_slide = (new GlfSlidePlayerItem)()
            new_slide.Slide = slide
            m_slides.Add slide & CStr(UBound(m_slides.Keys)), new_slide
            Set Slides = new_slide
        End Property

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

    Public Property Get RelativeTime() : RelativeTime = m_relTime: End Property
    Public Property Let RelativeTime(input) : m_relTime = input: End Property
    Public Property Get AbsoluteTime() : AbsoluteTime = m_absTime: End Property
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
        Set m_shows = CreateObject("Scripting.Dictionary")
        Set m_dofs = CreateObject("Scripting.Dictionary")
        Set m_slides = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        If UBound(m_lights) > -1 Then
            yaml = yaml & "  lights:" & vbCrLf
            Dim light
            For Each light in m_lights
                Dim light_parts
                light_parts = Split(light, "|")
                If UBound(light_parts) = 1 Then
                    yaml = yaml & "    " & light_parts(0) & ": ffffff%" & light_parts(1) & vbCrLf
                Else
                    yaml = yaml & "    " & light_parts(0) & ": " & light_parts(2) & "%" & light_parts(1) & vbCrLf
                End If
            Next
        End If
        ToYaml = yaml
    End Function

End Class


Class GlfSlidePlayer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "slide_player" : End Property

    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property   
    Public Property Get EventName(name)
        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_slide : Set new_slide = (new GlfSlidePlayerItem)()
        m_eventValues.Add newEvent.Raw, new_slide
        Set EventName = new_slide
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "slide_player_" & mode.name
        m_mode = mode.ModeName
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "slide_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_slide_player_play", "SlidePlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_slide_player_play"
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            'Fire Slide
            If useBcp = True Then
                bcpController.PlaySlide m_eventValues(evt).Slide, m_mode, m_events(evt).EventName, m_priority+m_eventValues(evt).Priority
            End If
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_slide_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim key
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_eventValues(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function SlidePlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim slide_player : Set slide_player = ownProps(1)
    Select Case evt
        Case "activate"
            slide_player.Activate
        Case "deactivate"
            slide_player.Deactivate
        Case "play"
            slide_player.Play(ownProps(2))
    End Select
    If IsObject(args(1)) Then
        Set SlidePlayerEventHandler = kwargs
    Else
        SlidePlayerEventHandler = kwargs
    End If
End Function

Class GlfSlidePlayerItem
	Private m_slide, m_action, m_expire, m_max_queue_time, m_method, m_priority, m_target, m_tokens
    
    Public Property Get Slide(): Slide = m_slide: End Property
    Public Property Let Slide(input)
        m_slide = input
    End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        m_action = input
    End Property

    Public Property Get Expire(): Expire = m_expire: End Property
    Public Property Let Expire(input)
        m_expire = input
    End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input)
        m_priority = input
    End Property

	Public default Function init()
        m_action = "play"
        m_slide = Empty
        m_expire = Empty
        m_priority = 0
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    "& m_slide & ":" & vbCrLf
        yaml = yaml & "      action: " & m_action & vbCrLf
        If Not IsEmpty(m_expire) Then
            yaml = yaml & "      expire: " & m_expire & "ms" & vbCrLf
        End If
        If m_priority <> 0 Then
            yaml = yaml & "      priority: " & m_priority & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class


Class GlfSoundPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_eventValues
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "sound_player" : End Property
    Public Property Get EventSounds() : EventSounds = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_sound : Set new_sound = (new GlfSoundPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_sound
        Set EventName = new_sound
        
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "sound_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "sound_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_sound_player_play", "SoundPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_sound_player_play"
            PlayOff evt
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            If m_eventValues(evt).Action = "stop" Then
                PlayOff evt
            Else
                glf_sound_buses(m_eventValues(evt).Sound.Bus).Play m_eventValues(evt)
            End If
        End If
    End Function

    Public Sub PlayOff(evt)
        glf_sound_buses(m_eventValues(evt).Sound.Bus).StopSoundWithKey m_eventValues(evt).Sound.File
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
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

Function SoundPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim sound_player : Set sound_player = ownProps(1)
    Select Case evt
        Case "activate"
            sound_player.Activate
        Case "deactivate"
            sound_player.Deactivate
        Case "play"
            'Dim block_queue
            sound_player.Play(ownProps(2))
            'If Not IsEmpty(block_queue) Then
            '    kwargs.Add "wait_for", block_queue
            'End If
    End Select
    If IsObject(args(1)) Then
        Set SoundPlayerEventHandler = kwargs
    Else
        SoundPlayerEventHandler = kwargs
    End If
End Function


Class GlfSoundPlayerItem
	Private m_sound, m_action, m_key, m_volume, m_loops
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property

    Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Sound()
        If IsNull(m_sound) Then
            Sound = Null
        Else
            Set Sound = m_sound
        End If
    End Property
	Public Property Let Sound(input)
        If glf_sounds.Exists(input) Then
            Set m_sound = glf_sounds(input)
        End If
    End Property
  
	Public default Function init()
        m_action = "play"
        m_sound = Null
        m_key = Empty
        m_volume = Empty
        m_loops = Empty
        Set Init = Me
	End Function

End Class

Class GlfStateMachine
    Private m_name
    Private m_player_var_name
    Private m_mode
    Private m_debug
    Private m_priority
    Private m_states
    Private m_transitions
    private m_base_device
 
    Private m_state
    Private m_persist_state
    Private m_starting_state
 
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
            Case "state":
                GetValue = State()
        End Select
    End Property

    Public Property Get State()
        If m_persist_state = True Then
            Dim s : s = GetPlayerState(m_player_var_name)
            If s=False Then
                State = Null
            Else
                State = s
            End If
        Else
            State = m_state
        End If
    End Property

    Public Property Let State(value)
        If m_persist_state = True Then
            SetPlayerState m_player_var_name, value
            m_state = value
        Else
            m_state = value
        End If
    End Property
    
    Public Property Get States(name)
        If m_states.Exists(name) Then
            Set States = m_states(name)
        Else
            Dim new_state : Set new_state = (new GlfStateMachineState)(name)
            m_states.Add name, new_state
            Set States = new_state
        End If
    End Property
    Public Property Get StateItems(): StateItems = m_states.Items(): End Property
 
    Public Property Get Transitions()
        Dim count : count = UBound(m_transitions.Keys)
        Dim new_transition : Set new_transition = (new GlfStateMachineTranistion)()
        m_transitions.Add CStr(count), new_transition
        Set Transitions = new_transition
    End Property
 
    Public Property Get PersistState(): PersistState = m_persist_state : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property

    Public Property Get StartingState(): StartingState = m_starting_state : End Property
    Public Property Let StartingState(value) : m_starting_state = value : End Property
 
    Public default Function init(name, mode)
        m_name = name
        m_player_var_name = "state_machine_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        m_persist_state = False
        m_starting_state = "start"
        m_state = Null
 
        Set m_states = CreateObject("Scripting.Dictionary")
        Set m_transitions = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "state_machine", Me)
        glf_state_machines.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        If IsNull(State()) Then
            StartState m_starting_state
        Else
            AddHandlersForCurrentState()
            RunShowForCurrentState()
        End If

    End Sub

    Public Sub Disable()
        RemoveHandlers()
        StopShowForCurrentState()
        m_state = Null
    End Sub

    Public Sub StartState(start_state)
        Log("Starting state " & start_state)
        If Not m_states.Exists(start_state) Then
            Log("Invalid state " & start_state)
            Exit Sub
        End If

        Dim state_config : Set state_config = m_states(start_state)

        State() = start_state
        If UBound(state_config.EventsWhenStarted().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStarted().Items()
                If evt.Evaluate() = True Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        AddHandlersForCurrentState()
        RunShowForCurrentState()
    End Sub

    Public Sub StopCurrentState()
        Log "Stopping state " & State()
        RemoveHandlers()
        If IsNull(state) Then
            Exit Sub
        End If
        Dim state_config : Set state_config = m_states(state)

        If UBound(state_config.EventsWhenStopped().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStopped().Items()
                If evt.Evaluate() = True Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        StopShowForCurrentState()

        State() = Null
    End Sub

    Public Sub RunShowForCurrentState()
        If IsNull(state) Then
            Exit Sub
        End If
        Log state
        Dim state_config : Set state_config = m_states(state)
        If Not IsNull(state_config.ShowWhenActive().Show) Then
            Dim show : Set show = state_config.ShowWhenActive
            Log "Starting show %s" & m_name & "_" & show.Key
            Dim new_running_show
            Set new_running_show = (new GlfRunningShow)(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key, show.Key, show, m_priority, Null, state_config.InternalCacheId)
        End If
    End Sub

    Public Sub StopShowForCurrentState()
        If IsNull(state) Then
            Exit Sub
        End If
        Dim state_config : Set state_config = m_states(state)
        If Not IsNull(state_config.ShowWhenActive().Show) Then
            Dim show : Set show = state_config.ShowWhenActive
            Log "Stopping show %s" & m_name & "_" & show.Key
            If glf_running_shows.Exists(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key) Then 
                glf_running_shows(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key).StopRunningShow()
            End If
        End If
    End Sub

    Public Sub AddHandlersForCurrentState()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            If transition.Source.Exists(State()) Then
                For Each evt in transition.Events.Items()
                    AddPinEventListener evt.EventName, m_name & "_" & transition.Target & "_" & evt.EventName & "_transition", "StateMachineTransitionHandler", m_priority+evt.Priority, Array("transition", Me, evt, transition)
                Next
            End If
        Next
    End Sub

    Public Sub RemoveHandlers()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            For Each evt in transition.Events.Items()
                RemovePinEventListener evt.EventName, m_name & "_" & transition.Target & "_" & evt.EventName & "_transition"
            Next
        Next
    End Sub

    Public Sub MakeTransition(transition)

        Log "Transitioning from " & State() & " to " & transition.Target
        StopCurrentState()
        If UBound(transition.EventsWhenTransitioning().Keys()) > -1 Then
            Dim evt
            For Each evt in transition.EventsWhenTransitioning().Items()
                If evt.Evaluate() = True Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If
        
        StartState transition.Target

    End Sub

    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class
 
Class GlfStateMachineState
	Private m_name, m_label, m_show_when_active, m_events_when_started, m_events_when_stopped, m_internal_cache_id
 

    Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property

	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property
 
    Public Property Get Label(): Label = m_label: End Property
    Public Property Let Label(input): m_label = input: End Property
 
    Public Property Get ShowWhenActive()
        Set ShowWhenActive = m_show_when_active
    End Property

    Public Property Get EventsWhenStarted(): Set EventsWhenStarted = m_events_when_started: End Property
    Public Property Let EventsWhenStarted(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_started.Add newEvent.Raw, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenStopped(): Set EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(input)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_stopped.Add newEvent.Raw, newEvent
        Next
    End Property
 
	Public default Function init(name)
        m_name = name
        m_label = Empty
        m_internal_cache_id = -1
        Set m_show_when_active = (new GlfShowPlayerItem)()
        Set m_events_when_started = CreateObject("Scripting.Dictionary")
        Set m_events_when_stopped = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class
 
Class GlfStateMachineTranistion
	Private m_name, m_sources, m_target, m_events, m_events_when_transitioning

    Public Property Get Source(): Set Source = m_sources: End Property
    Public Property Let Source(value)
        Dim x
        For x=0 to UBound(value)
            m_sources.Add value(x), True
        Next    
    End Property
 
    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input): m_target = input: End Property  

    Public Property Get Events(): Set Events = m_events: End Property
    Public Property Let Events(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events.Add newEvent.Raw, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenTransitioning(): Set EventsWhenTransitioning = m_events_when_transitioning: End Property
    Public Property Let EventsWhenTransitioning(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_transitioning.Add newEvent.Raw, newEvent
        Next    
    End Property
 
	Public default Function init()
        Set m_sources = CreateObject("Scripting.Dictionary")
        m_target = Empty
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_events_when_transitioning = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class


Public Function StateMachineTransitionHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim state_machine : Set state_machine = ownProps(1)
    Select Case evt
        Case "transition"
            Dim glf_event : Set glf_event = ownProps(2)
            If glf_event.Evaluate() = True Then
                state_machine.MakeTransition ownProps(3)
            Else
                If glf_debug_level = "Debug" Then
                    glf_debugLog.WriteToLog "State machine transition",  "failed condition: " & glf_event.Raw
                End If
            End If
    End Select
    If IsObject(args(1)) Then
        Set StateMachineTransitionHandler = kwargs
    Else
        StateMachineTransitionHandler = kwargs
    End If
End Function
Class GlfTilt

    Private m_name
    Private m_priority
    Private m_base_device
    Private m_reset_warnings_events
    Private m_tilt_events
    Private m_tilt_warning_events
    Private m_tilt_slam_tilt_events
    Private m_settle_time
    Private m_warnings_to_tilt
    Private m_multiple_hit_window
    Private m_tilt_warning_switch
    Private m_tilt_switch
    Private m_slam_tilt_switch
    Private m_last_tilt_warning_switch 
    Private m_last_warning
    Private m_balls_to_collect
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "tilt_settle_ms_remaining":
                GetValue = TiltSettleMsRemaining()
            Case "tilt_warnings_remaining":
                GetValue = TiltWarningsRemaining()
        End Select
    End Property

    Public Property Let ResetWarningEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_warnings_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let SettleTime(value): Set m_settle_time = CreateGlfInput(value): End Property
    Public Property Let WarningsToTilt(value): Set m_warnings_to_tilt = CreateGlfInput(value): End Property
    Public Property Let MultipleHitWindow(value): Set m_multiple_hit_window = CreateGlfInput(value): End Property

    Private Property Get TiltSettleMsRemaining()
        TiltSettleMsRemaining = 0
        If m_last_tilt_warning_switch > 0 Then
            Dim delta
            delta = m_settle_time.Value - (gametime - m_last_tilt_warning_switch)
            If delta > 0 Then
                TiltSettleMsRemaining = delta
            End If
        End If
    End Property

    Private Property Get TiltWarningsRemaining() 
        TiltWarningsRemaining = 0

        If glf_gameStarted Then
            TiltWarningsRemaining = m_warnings_to_tilt.Value() - GetPlayerState("tilt_warnings")
        End If   
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(mode)
        m_name = "tilt_" & mode.name
        m_priority = mode.Priority
        Set m_reset_warnings_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_warning_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_slam_tilt_events = CreateObject("Scripting.Dictionary")
        Set m_settle_time = CreateGlfInput(0)
        Set m_warnings_to_tilt = CreateGlfInput(0)
        Set m_multiple_hit_window = CreateGlfInput(0)
        m_tilt_switch = Empty
        m_tilt_warning_switch = Empty
        m_slam_tilt_switch = Empty
        m_last_tilt_warning_switch = 0
        m_last_warning = 0
        m_balls_to_collect = 0
        Set m_base_device = (new GlfBaseModeDevice)(mode, "tilt", Me)
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_reset_warnings_events.Keys()
            AddPinEventListener m_reset_warnings_events(evt).EventName, m_name & "_" & evt & "_reset_warnings", "TiltHandler", m_priority+m_reset_warnings_events(evt).Priority, Array("reset_warnings", Me, m_reset_warnings_events(evt))
        Next
        For Each evt in m_tilt_events.Keys()
            AddPinEventListener m_tilt_events(evt).EventName, m_name & "_" & evt & "_tilt", "TiltHandler", m_priority+m_tilt_events(evt).Priority, Array("tilt", Me, m_tilt_events(evt))
        Next
        For Each evt in m_tilt_warning_events.Keys()
            AddPinEventListener m_tilt_warning_events(evt).EventName, m_name & "_" & evt & "_tilt_warning", "TiltHandler", m_priority+m_tilt_warning_events(evt).Priority, Array("tilt_warning", Me, m_tilt_warning_events(evt))
        Next
        For Each evt in m_tilt_slam_tilt_events.Keys()
            AddPinEventListener m_tilt_slam_tilt_events(evt).EventName, m_name & "_" & evt & "_slam_tilt", "TiltHandler", m_priority+m_tilt_slam_tilt_events(evt).Priority, Array("slam_tilt", Me, m_tilt_slam_tilt_events(evt))
        Next
        
        AddPinEventListener  "s_tilt_warning_active", m_name & "_tilt_warning_switch_active", "TiltHandler", m_priority, Array("_tilt_warning_switch_active", Me)
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_reset_warnings_events.Keys()
            RemovePinEventListener m_reset_warnings_events(evt).EventName, m_name & "_" & evt & "_reset_warnings"
        Next
        For Each evt in m_tilt_events.Keys()
            RemovePinEventListener m_tilt_events(evt).EventName, m_name & "_" & evt & "_tilt"
        Next
        For Each evt in m_tilt_warning_events.Keys()
            RemovePinEventListener m_tilt_warning_events(evt).EventName, m_name & "_" & evt & "_tilt_warning"
        Next
        For Each evt in m_tilt_slam_tilt_events.Keys()
            RemovePinEventListener m_tilt_slam_tilt_events(evt).EventName, m_name & "_" & evt & "_slam__tilt"
        Next

        RemovePinEventListener "s_tilt_warning_active", m_name & "_tilt_warning_switch_active"

    End Sub

    Public Sub TiltWarning()
        'Process a tilt warning.
        'If the number of warnings is than the number to cause a tilt, a tilt will be
        'processed.

        m_last_tilt_warning_switch = gametime
        If glf_gameStarted = False Or glf_gameTilted = True Then
            Exit Sub
        End If
        Log "Tilt Warning"
        m_last_warning = gametime
        SetPlayerState "tilt_warnings", GetPlayerState("tilt_warnings")+1
        Dim warnings : warnings = GetPlayerState("tilt_warnings")
        Dim warnings_to_tilt : warnings_to_tilt = m_warnings_to_tilt.Value()
        If warnings>=warnings_to_tilt Then
            Tilt()
        Else
            Dim kwargs
            Set kwargs = GlfKwargs()
            With kwargs
                .Add "warnings", warnings
                .Add "warnings_remaining", warnings_to_tilt - warnings
            End With
            DispatchPinEvent "tilt_warning", kwargs
            DispatchPinEvent "tilt_warning_" & warnings, Null
        End If
    End Sub

    Public Sub ResetWarnings()
        'Reset the tilt warnings for the current player.
        If glf_gamestarted = False or glf_gameEnding = True Then
            Exit Sub
        End If
        SetPlayerState "tilt_warnings", 0
    End Sub

    Public Sub Tilt()
        'Cause the ball to tilt.
        'This will post an event called *tilt*, set the game mode's ``tilted``
        'attribute to *True*, disable the flippers and autofire devices, end the
        'current ball, and wait for all the balls to drain.
        If glf_gameStarted = False or glf_gameTilted=True or glf_gameEnding = True Then
            Exit Sub
        End If
        glf_gametilted = True
        m_balls_to_collect = glf_BIP
        Log "Processing Tilt. Balls to collect: " & m_balls_to_collect
        DispatchPinEvent "tilt", Null
        AddPinEventListener GLF_BALL_ENDING, m_name & "_ball_ending", "TiltHandler", 20, Array("tilt_ball_ending", Me)
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "TiltHandler", 999999, Array("tilt_ball_drain", Me)
        Glf_EndBall()
    End Sub

    Public Sub TiltedBallDrain(unclaimed_balls)
        Log "Tilted ball drain, unclaimed balls: " & unclaimed_balls
        m_balls_to_collect = m_balls_to_collect - unclaimed_balls
        Log "Tilted ball drain. Balls to collect: " & m_balls_to_collect
        If m_balls_to_collect <= 0 Then
            TiltDone()
        End If
    End Sub

    Public Sub HandleTiltWarningSwitch()
        Log "Handling Tilt Warning Switch"
        If m_last_warning = 0 Or (m_last_warning + m_multiple_hit_window.Value()) <= gametime Then
            TiltWarning()
        End If
    End Sub

    Public Sub BallEndingTilted()
        If m_balls_to_collect<=0 Then
            TiltDone()
        End If
    End Sub

    Public Sub TiltDone()
        If TiltSettleMsRemaining() > 0 Then
            SetDelay "delay_tilt_clear", "TiltHandler", Array(Array("tilt_done", Me), Null), TiltSettleMsRemaining()
            Exit Sub
        End If
        Log "Tilt Done"
        RemovePinEventListener GLF_BALL_ENDING, m_name & "_ball_ending"
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        glf_gameTilted = False
        DispatchPinEvent m_name & "_clear", Null
    End Sub
    
    Public Sub SlamTilt()
        'Process a slam tilt.
        'This method posts the *slam_tilt* event and (if a game is active) sets
        'the game mode's ``slam_tilted`` attribute to *True*.
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TiltHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim tilt : Set tilt = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set TiltHandler = kwargs
                Else
                    TiltHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "_tilt_warning_switch_active":
            tilt.HandleTiltWarningSwitch
        Case "tilt_ball_ending"
            kwargs.Add "wait_for", tilt.Name & "_clear"
            tilt.BallEndingTilted
        Case "tilt_ball_drain"
            tilt.TiltedBallDrain kwargs
            kwargs = kwargs -1
        Case "tilt_done"
            tilt.TiltDone
        Case "reset_warnings"
            tilt.ResetWarnings
    End Select

    If IsObject(args(1)) Then
        Set TiltHandler = kwargs
    Else
        TiltHandler = kwargs
    End If
End Function

Class GlfTimedSwitches

    Private m_name
    Private m_priority
    Private m_base_device
    Private m_time
    Private m_switches
    Private m_events_when_active
    Private m_events_when_released
    Private m_active_switches
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        'Select Case value
            'Case   
        'End Select
        GetValue = True
    End Property

    Public Property Let Time(value): Set m_time = CreateGlfInput(value): End Property
    Public Property Get Time(): Time = m_time.Value(): End Property
    
    Public Property Let Switches(value): m_switches = value: End Property
    
    Public Property Let EventsWhenActive(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_active.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenReleased(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_released.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(name, mode)
        m_name = "timed_switch_" & name
        m_priority = mode.Priority
        Set m_events_when_active = CreateObject("Scripting.Dictionary")
        Set m_events_when_released = CreateObject("Scripting.Dictionary")
        Set m_time = CreateGlfInput(0)
        m_switches = Array()
        Set m_active_switches = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "timed_switch", Me)
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim switch
        For Each switch in m_switches
            AddPinEventListener switch & "_active", m_name & "_active", "TimedSwitchHandler", m_priority, Array("active", Me, switch)
            AddPinEventListener switch & "_inactive", m_name & "_inactive", "TimedSwitchHandler", m_priority, Array("inactive", Me, switch)
        Next
    End Sub

    Public Sub Deactivate()
        Dim switch
        For Each switch in m_switches
            RemovePinEventListener switch & "_active", m_name & "_active"
            RemovePinEventListener switch & "_inactive", m_name & "_inactive"
        Next
    End Sub

    Public Sub SwitchActive(switch)
        If UBound(m_active_switches.Keys()) = -1 Then
            Dim evt
            For Each evt in m_events_when_active.Keys()
                Log "Switch Active: " & switch & ". Event: " & m_events_when_active(evt).EventName
                DispatchPinEvent m_events_when_active(evt).EventName, Null
            Next
        End If
        If Not m_active_switches.Exists(switch) Then
            m_active_switches.Add switch, True
        End If
    End Sub

    Public Sub SwitchInactive(switch)
        RemoveDelay m_name & "_" & switch & "_active"
        If m_active_switches.Exists(switch) Then
            m_active_switches.Remove switch
            If UBound(m_active_switches.Keys()) = -1 Then
                Dim evt
                For Each evt in m_events_when_released.Keys()
                    Log "Switch Release: " & switch & ". Event: " & m_events_when_released(evt).EventName
                    DispatchPinEvent m_events_when_released(evt).EventName, Null
                Next
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TimedSwitchHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim timed_switch : Set timed_switch = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set TimedSwitchHandler = kwargs
                Else
                    TimedSwitchHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "active"
            SetDelay timed_switch.Name & "_" & ownProps(2) & "_active" , "TimedSwitchHandler" , Array(Array("passed_time", timed_switch, ownProps(2)),Null), timed_switch.Time
        Case "passed_time"
            timed_switch.SwitchActive ownProps(2)
        Case "inactive"
            timed_switch.SwitchInactive ownProps(2)
    End Select

    If IsObject(args(1)) Then
        Set TimedSwitchHandler = kwargs
    Else
        TimedSwitchHandler = kwargs
    End If
End Function



Class GlfTimer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

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
    Private m_restart_on_complete
    Private m_start_running

    Public Property Get Name() : Name = m_name : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    

    Public Property Get ControlEvents()
        Dim count : count = UBound(m_control_events.Keys) 
        Dim newEvent : Set newEvent = (new GlfTimerControlEvent)()
        m_control_events.Add CStr(count), newEvent
        Set ControlEvents = newEvent
    End Property
    Public Property Get StartValue() : StartValue = m_start_value : End Property
    Public Property Get EndValue() : EndValue = m_end_value : End Property
    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let StartValue(value) : Set m_start_value = CreateGlfInput(value) : End Property
    Public Property Let EndValue(value) : Set m_end_value = CreateGlfInput(value) : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let MaxValue(value) : m_max_value = value : End Property
    Public Property Let RestartOnComplete(value) : m_restart_on_complete = value : End Property
    Public Property Let StartRunning(value) : m_start_running = value : End Property
    Public Property Let TickInterval(value)
        m_tick_interval = value
        m_starting_tick_interval = value
    End Property

    Public Property Get GetValue(value)
        Select Case value
            Case "ticks"
                GetValue = m_ticks
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
        m_starting_tick_interval = 1000
        m_restart_on_complete = False
        m_start_running = False
        Set m_start_value = CreateGlfInput(0)
        Set m_end_value = CreateGlfInput(-1)

        Set m_control_events = CreateObject("Scripting.Dictionary")
        m_running = False

        Set m_base_device = (new GlfBaseModeDevice)(mode, "timer", Me)

        glf_timers.Add name, Me

        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_control_events.Keys
            AddPinEventListener m_control_events(evt).EventName.EventName, m_name & "_action", "TimerEventHandler", m_priority+m_control_events(evt).EventName.Priority, Array("action", Me, m_control_events(evt))
        Next
        m_ticks = m_start_value.Value
        m_ticks_remaining = m_ticks
        Log "Activating Timer"
        If m_start_running = True Then
            StartTimer()
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_control_events.Keys
            RemovePinEventListener m_control_events(evt).EventName.EventName, m_name & "_action"
        Next
        RemoveDelay m_name & "_tick"
        m_running = False
    End Sub

    Public Sub Action(controlEvent)

        If Not IsNull(controlEvent.EventName) Then
            If controlEvent.EventName().Evaluate() = False Then
                Exit Sub
            End IF
        End If

        dim value : value = controlEvent.Value
        Select Case controlEvent.Action
            Case "add"
                Add value
            Case "subtract"
                Subtract value
            Case "jump"
                Jump value
            Case "start"
                StartTimer()
            Case "stop"
                StopTimer()
            Case "reset"
                Reset()
            Case "restart"
                Restart()
            Case "pause"
                Pause value
            Case "set_tick_interval"
                SetTickInterval value
            Case "change_tick_interval"
                ChangeTickInterval value
            Case "reset_tick_interval"
                SetTickInterval m_starting_tick_interval
        End Select

    End Sub

    Private Sub StartTimer()
        If m_running Then
            Exit Sub
        End If

        Log "Starting Timer"
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
        Log "Stopping Timer"
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
        Log "Pausing Timer for "&pause_ms&" ms"
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
        Log "Timer Tick"
        If Not m_running Then
            Log "Timer is not running. Will remove."
            Exit Sub
        End If

        Dim newValue
        If m_direction = "down" Then
            newValue = m_ticks - 1
        Else
            newValue = m_ticks + 1
        End If
        
        Log "ticking: old value: "& m_ticks & ", new Value: " & newValue & ", target: "& m_end_value.Value
        m_ticks = newValue
        If Not PostTickEvents() Then
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval    
        End If
    End Sub

    Private Function CheckForDone

        ' Checks to see if this timer is done. Automatically called anytime the
        ' timer's value changes.
        Log "Checking to see if timer is done. Ticks: "&m_ticks&", End Value: "&m_end_value.Value&", Direction: "& m_direction

        if m_direction = "up" And m_end_value.Value<>-1 And m_ticks >= m_end_value.Value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If m_direction = "down" And m_ticks <= m_end_value.Value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If m_end_value.Value<>-1 Then 
            m_ticks_remaining = abs(m_end_value.Value - m_ticks)
        End If
        Log "Timer is not done"

        CheckForDone = False

    End Function

    Private Sub TimerComplete

        Log "Timer Complete"

        StopTimer()
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        
        If m_restart_on_complete Then
            Log "Restart on complete: True"
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
        Log "Resetting timer. New value: "& m_start_value.Value
        Jump m_start_value.Value
    End Sub

    Private Sub Jump(timer_value)
        m_ticks = timer_value

        If m_max_value and m_ticks > m_max_value Then
            m_ticks = m_max_value
        End If

        CheckForDone()
    End Sub

    Public Sub ChangeTickInterval(change)
        m_tick_interval = m_tick_interval * change
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
            Log "Ticks: "&m_ticks&", Remaining: " & m_ticks_remaining
        End If
    End Function

    Public Sub Add(timer_value) 
        Dim new_value

        new_value = m_ticks + timer_value

        If Not IsEmpty(m_max_value) And new_value > m_max_value Then
            new_value = m_max_value
        End If
        m_ticks = new_value
        timer_value = new_value - timer_value

        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_added", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_added", kwargs
        CheckForDone()
    End Sub

    Public Sub Subtract(timer_value)
        m_ticks = m_ticks - timer_value
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_subtracted", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_subtracted", kwargs
        
        CheckForDone()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    'debug.print "TimerEventHandler: " & timer.Name & ": " & evt
    Select Case evt
        Case "action"
            Dim controlEvent : Set controlEvent = ownProps(2)
            timer.Action controlEvent
        Case "tick"
            timer.Tick 
    End Select
    If IsObject(args(1)) Then
        Set TimerEventHandler = kwargs
    Else
        TimerEventHandler = kwargs
    End If
End Function

Class GlfTimerControlEvent
	Private m_event, m_action, m_value
  
	Public Property Get EventName(): Set EventName = m_event: End Property
    Public Property Let EventName(input)
        Dim newEvent : Set newEvent = (new GlfEvent)(input)
        Set m_event = newEvent
    End Property

    Public Property Get Action(): Action = m_action : End Property
    Public Property Let Action(input): m_action = input : End Property

    Public Property Get Value()
        If Not IsNull(m_value) Then
            Value = m_value.Value
        Else
            Value = 0
        End If
    End Property
    Public Property Let Value(input)
        Set m_value = CreateGlfInput(input)
    End Property

	Public default Function init()
        m_event = Null
        m_action = Empty
        m_value = Null
	    Set Init = Me
	End Function

End Class
Class GlfVariablePlayer

    Private m_priority
    Private m_name
    Private m_mode
    Private m_events
    Private m_debug
    private m_base_device

    Private m_value

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property 
    Public Property Get EventName(name)
        Dim newEvent : Set newEvent = (new GlfVariablePlayerEvent)(name)
        m_events.Add newEvent.BaseEvent.Raw, newEvent
        Set EventName = newEvent
    End Property
   
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "variable_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "variable_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_" & evt & "_variable_player_play", "VariablePlayerEventHandler", m_priority+m_events(evt).BaseEvent.Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_" & evt & "_variable_player_play"
        Next
    End Sub

    Public Sub Play(evt)
        Log "Playing: " & evt
        If m_events(evt).BaseEvent.Evaluate() = False Then
            Exit Sub
        End If
        Dim vKey, v
        For Each vKey in m_events(evt).Variables.Keys
            Set v = m_events(evt).Variable(vKey)
            Dim varValue : varValue = v.VariableValue
            Select Case v.Action
                Case "add"
                    Log "Add Variable " & vKey & ". New Value: " & CStr(GetPlayerState(vKey) + varValue) & " Old Value: " & CStr(GetPlayerState(vKey))
                    SetPlayerState vKey, GetPlayerState(vKey) + varValue
                Case "add_machine"
                    Log "Add Machine Variable " & vKey & ". New Value: " & CStr(GetPlayerState(vKey) + varValue) & " Old Value: " & CStr(GetPlayerState(vKey))
                    'SetPlayerState vKey, GetPlayerState(vKey) + varValue
                    glf_machine_vars(vkey).Value = glf_machine_vars(vkey).Value + varValue
                Case "set"
                    Log "Setting Variable " & vKey & ". New Value: " & CStr(varValue)
                    SetPlayerState vKey, varValue
                Case "set_machine"
                    Log "Setting Machine Variable " & vKey & ". New Value: " & CStr(varValue)
                    glf_machine_vars(vkey).Value = varValue
        End Select
        Next
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim key
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_variable_player", message
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

    Public Function ToYaml()
        Dim yaml
        Dim key
        If UBound(m_variables.Keys) > -1 Then
            For Each key in m_variables.keys
                yaml = yaml & "    " & key & ": " & vbCrLf
                yaml = yaml & m_variables(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfVariablePlayerItem
	Private m_block, m_show, m_float, m_int, m_string, m_player, m_action, m_type
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Block(): Block = m_block End Property
    Public Property Let Block(input): m_block = input End Property

	Public Property Let Float(input): Set m_float = CreateGlfInput(input): m_type = "float" : End Property
  
	Public Property Let Int(input): Set m_int = CreateGlfInput(input): m_type = "int" : End Property
  
	Public Property Let String(input) : Set m_string = CreateGlfInput(input) : m_type = "string" : End Property

    Public Property Get VariableType(): VariableType = m_type: End Property
    Public Property Get VariableValue()
        Select Case m_type
            Case "float"
                VariableValue = m_float.Value()
            Case "int"
                VariableValue = m_int.Value()
            Case "string"
                VariableValue = m_string.Value()
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

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "      "& "action" & ": " & m_action & vbCrLf
        If m_block = True Then
            yaml = yaml & "      "& "block" & ": true" & vbCrLf
        End If
        If Not IsEmpty(m_int) Then
            yaml = yaml & "      "& "int" & ": " & m_int.Raw & vbCrLf
        End If
        If Not IsEmpty(m_float) Then
            yaml = yaml & "      "& "float" & ": " & m_float.Raw & vbCrLf
        End If
        If Not IsEmpty(m_string) Then
            yaml = yaml & "      "& "string" & ": " & m_string.Raw & vbCrLf
        End If
        If Not IsEmpty(m_player) Then
            yaml = yaml & "      "& "player" & ": " & m_strinm_playerg & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class


Function CreateMachineVar(name)
	Dim machine_var : Set machine_var = (new GlfMachineVars)(name)
	Set CreateMachineVar = machine_var
End Function

Class GlfMachineVars

    Private m_name
	Private m_persist
    Private m_value_type
    Private m_value

    Public Property Get Name(): Name = m_name : End Property
    Public Property Let Name(input): m_name = input : End Property

    Public Property Get Value(): Value = m_value : End Property
    Public Property Let Value(input): m_value = input : End Property

    Public Property Let InitialValue(input)
        m_value = input
    End Property

    Public Property Get Persist(): Persist = m_persist : End Property
    Public Property Let Persist(input): m_persist = input : End Property

    Public Property Get ValueType(): ValueType = m_value_type : End Property
    Public Property Let ValueType(input): m_value_type = input : End Property

    Public Function GetValue()
        GetValue = m_value
    End Function

	Public default Function init(name)
        m_name = name
        m_persist = True
        m_value_type = "int"
        m_value = 0
        glf_machine_vars.Add name, Me
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
    
    RemoveDelay name
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
    'Glf_WriteDebugLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc & ", ExecutionTime: " & executionTime
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
                'Glf_WriteDebugLog "Delay", "Removing delay for " & name & " and  Execution Time: " & delayQueueMap(name)
                delayQueue(delayQueueMap(name)).Remove name
            End If
            delayQueueMap.Remove name
            RemoveDelay = True
            'Glf_WriteDebugLog "Delay", "Removing delay for " & name
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
                            'Glf_WriteDebugLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
                            GetRef(delayObject.Callback)(delayObject.Args)    
                End If
            Next
            delayQueue.Remove queueItem
        End If
    Next
End Sub

Function CreateGlfAutoFireDevice(name)
	Dim flipper : Set flipper = (new GlfAutoFireDevice)(name)
	Set CreateGlfAutoFireDevice = flipper
End Function

Class GlfAutoFireDevice

    Private m_name
    Private m_device_name
    Private m_enable_events
    Private m_disable_events
    Private m_enabled
    Private m_switch
    Private m_action_cb
    Private m_disabled_cb
    Private m_enabled_cb
    Private m_exclude_from_ball_search
    Private m_debug

    Public Property Let Switch(value)
        m_switch = value
    End Property
    Public Property Let ExcludeFromBallSearch(value) : m_exclude_from_ball_search = value : End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let DisabledCallback(value) : m_disabled_cb = value : End Property
    Public Property Let EnabledCallback(value) : m_enabled_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "AutoFireDeviceEventHandler", 1000, Array("enable", Me)
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
            AddPinEventListener evt, m_name & "_disable", "AutoFireDeviceEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "auto_fire_coil_" & name
        m_device_name = name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_enabled = False
        m_action_cb = Empty
        m_disabled_cb = Empty
        m_enabled_cb = Empty
        m_switch = Empty
        m_debug = False
        m_exclude_from_ball_search = False
        glf_autofiredevices.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        If Not IsEmpty(m_enabled_cb) Then
            GetRef(m_enabled_cb)(Null)
        End If
        If Not IsEmpty(m_switch) Then
            AddPinEventListener m_switch & "_active", m_name & "_active", "AutoFireDeviceEventHandler", 1000, Array("activate", Me)
            AddPinEventListener m_switch & "_inactive", m_name & "_inactive", "AutoFireDeviceEventHandler", 1000, Array("deactivate", Me)
        End If
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        If Not IsEmpty(m_disabled_cb) Then
            GetRef(m_disabled_cb)(Null)
        End If
        Deactivate(Null)
        RemovePinEventListener m_switch & "_active", m_name & "_active"
        RemovePinEventListener m_switch & "_inactive", m_name & "_inactive"
    End Sub

    Public Sub Activate(active_ball)
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(1, active_ball))
        End If
        DispatchPinEvent m_name & "_activate", Null
    End Sub

    Public Sub Deactivate(active_ball)
        Log "Deactivating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(0, active_ball))
        End If
        DispatchPinEvent m_name & "_deactivate", Null
    End Sub

    Public Sub BallSearch(phase)
        If m_exclude_from_ball_search = True Then
            Exit Sub
        End If
        Log "Ball Search, phase " & phase
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(1, Null))
        End If
        SetDelay m_name & "ball_search_deactivate", "AutoFireDeviceEventHandler", Array(Array("deactivate", Me), Null), 150
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function AutoFireDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim flipper : Set flipper = ownProps(1)
    Select Case evt
        Case "enable"
            flipper.Enable
        Case "disable"
            flipper.Disable
        Case "activate"
            flipper.Activate kwargs
        Case "deactivate"
            flipper.Deactivate kwargs
    End Select
    If IsObject(args(1)) Then
        Set AutoFireDeviceEventHandler = kwargs
    Else
        AutoFireDeviceEventHandler = kwargs
    End If
End Function
Function CreateGlfBallDevice(name)
	Dim device : Set device = (new GlfBallDevice)(name)
	Set CreateGlfBallDevice = device
End Function

Class GlfBallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeout
    Private m_eject_enable_time
    Private m_balls
    Private m_balls_in_device
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_balls_to_eject
    Private m_ejecting_all
    Private m_ejecting
    Private m_mechanical_eject
    Private m_eject_targets
    Private m_entrance_count_delay
    Private m_incoming_balls
    Private m_lost_balls
    Private m_exclude_from_ball_search
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "balls":
                GetValue = m_balls_in_device
        End Select
    End Property
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

    Public Property Let AddIncomingBalls(value) : m_incoming_balls = m_incoming_balls + value : End Property
    Public Property Get IncomingBalls() : IncomingBalls = m_incoming_balls : End Property

    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    Public Property Let EjectEnableTime(value) : m_eject_enable_time = value : End Property
        
    Public Property Let EjectTimeout(value) : m_eject_timeout = value : End Property
    Public Property Let EntranceCountDelay(value) : m_entrance_count_delay = value : End Property
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
    Public Property Let ExcludeFromBallSearch(value) : m_exclude_from_ball_search = value : End Property

    Public Property Let Debug(value) : m_debug = value : End Property
        
	Public default Function init(name)
        m_name = "balldevice_" & name
        m_ball_switches = Array()
        m_eject_all_events = Array()
        m_eject_targets = Array()
        m_balls = Array()
        m_debug = False
        m_default_device = False
        m_ejecting = False
        m_eject_callback = Null
        m_ejecting_all = False
        m_balls_to_eject = 0
        m_balls_in_device = 0
        m_lost_balls = 0
        m_mechanical_eject = False
        m_eject_timeout = 1000
        m_eject_enable_time = 0
        m_entrance_count_delay = 500
        m_incoming_balls = 0
        m_exclude_from_ball_search = False
        glf_ball_devices.Add name, Me
	    Set Init = Me
	End Function

    Public Sub BallEntering(ball, switch)
        Log "Ball Entering" 
        If m_default_device = False Then
            SetDelay m_name & "_" & switch & "_ball_enter", "BallDeviceEventHandler", Array(Array("ball_enter", Me, switch), ball), m_entrance_count_delay
        Else
            BallEnter ball, switch
        End If
    End Sub

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_switch" & switch & "_eject_timeout"
        Set m_balls(switch) = ball
        m_balls_in_device = m_balls_in_device + 1
        Log "Ball Entered"
        If m_lost_balls > 0 Then
            m_lost_balls = m_lost_balls - 1
            Log "Lost Ball Found"
            If m_lost_balls = 0 Then
                RemoveDelay m_name & "_clear_lost_balls"
            End If
            Exit Sub
        End If

        Dim unclaimed_balls: unclaimed_balls = 1
        If m_incoming_balls > 0 Then
            unclaimed_balls = 0
        End If
        If unclaimed_balls > 0 Then
            unclaimed_balls = DispatchRelayPinEvent(m_name & "_ball_enter", 1)
        End If        
        Log "Unclaimed Balls: " & unclaimed_balls
        DispatchPinEvent m_name & "_ball_entered", Null
        If unclaimed_balls > 0 Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        RemoveDelay m_name & "_" & switch & "_ball_enter"
        If m_ejecting = False And m_mechanical_eject = False Then
            Log "Ball Lost, Wasn't Ejecting"
            m_lost_balls = m_lost_balls + 1
            SetDelay m_name & "_clear_lost_balls", "BallDeviceEventHandler", Array(Array("clear_lost_balls", Me), Null), 3000
        End If
        m_balls(switch) = Null
        m_balls_in_device = m_balls_in_device - 1
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechanical_eject = True And m_eject_timeout > 0 Then
            SetDelay m_name & "_switch" & switch & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), ball), m_eject_timeout
        End If
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        m_ejecting = False

        If m_incoming_balls > 0 Then
            m_incoming_balls = m_incoming_balls - 1
        End If
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
        
        If Not IsNull(m_eject_callback) Then
            If Not IsNull(m_balls(0)) Then
                Log "Ejecting."
                SetDelay m_name & "_switch0_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_balls(0)), m_eject_timeout
                m_ejecting = True
            
                GetRef(m_eject_callback)(m_balls(0))
                If m_eject_enable_time > 0 Then
                    SetDelay m_name & "_eject_enable_time", "BallDeviceEventHandler", Array(Array("eject_enable_complete", Me), m_balls(0)), m_eject_enable_time
                End If
            Else
                SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), Null), 600
            End If
        End If
    End Sub

    Public Sub EjectEnableComplete
        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(Null)
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

    Public Sub ClearLostBalls()
        m_lost_balls = 0
    End Sub

    Public Sub BallSearch(phase)
        If m_exclude_from_ball_search = True Then
            Exit Sub
        End If
        Log "Ball Search, phase " & phase
        If m_default_device = True Then
            Exit Sub
        End If
        If phase = 1 And HasBall() Then
            Exit Sub
        End If
        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(m_balls(0))
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml + "    ball_switches: " & Join(m_ball_switches, ",") & vbCrLf
        yaml = yaml + "    mechanical_eject: " & m_mechanical_eject & vbCrLf
        
        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallDeviceEventHandler(args)
    Dim ownProps, ball
    ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ballDevice : Set ballDevice = ownProps(1)
    Dim switch
    'debug.print "Ball Device: " & ballDevice.Name & ". Event: " & evt
    Select Case evt
        Case "ball_entering"
            Set ball = args(1)
            switch = ownProps(2)
            ballDevice.BallEntering ball, switch
        Case "ball_enter"
            Set ball = args(1)
            switch = ownProps(2)
            ballDevice.BallEnter ball, switch
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_eject_all"
            ballDevice.EjectAll
        Case "ball_exiting"
            switch = ownProps(2)
            If RemoveDelay(ballDevice.Name & "_" & switch & "_ball_enter") = False Then
                Set ball = args(1)
                ballDevice.BallExiting ball, switch
            End If
        Case "eject_timeout"
            Set ball = args(1)
            ballDevice.BallExitSuccess ball
        Case "eject_enable_complete"
            ballDevice.EjectEnableComplete
        Case "clear_lost_balls"
            ballDevice.ClearLostBalls
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
    Private m_enabled
    Private m_active
    Private m_ball_search_hold_time
    Private m_exclude_from_ball_search
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "active":
                GetValue = m_active
        End Select
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
    Public Property Let ActivateEvents(value) 
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_activate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let DeactivateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_deactivate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ActivationTime(value) : Set m_activation_time = CreateGlfInput(value) : End Property
    Public Property Let ActivationSwitches(value) : m_activation_switches = value : End Property
    Public Property Let BallSearchHoldTime(value) : Set m_ball_search_hold_time = CreateGlfInput(value) : End Property
    Public Property Let ExcludeFromBallSearch(value) : m_exclude_from_ball_search = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "diverter_" & name
        m_enable_events = Array()
        m_disable_events = Array()
        Set m_activate_events = CreateObject("Scripting.Dictionary")
        Set m_deactivate_events = CreateObject("Scripting.Dictionary")
        m_activation_switches = Array()
        Set m_activation_time = CreateGlfInput(0)
        Set m_ball_search_hold_time = CreateGlfInput(1000)
        m_debug = False
        m_enabled = False
        m_active = False
        m_exclude_from_ball_search = False
        glf_diverters.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_activate_events.Keys()
            AddPinEventListener m_activate_events(evt).EventName, m_name & "_" & evt & "_activate", "DiverterEventHandler", 1000, Array("activate", Me, m_activate_events(evt))
        Next
        For Each evt in m_deactivate_events.Keys()
            AddPinEventListener m_deactivate_events(evt).EventName, m_name & "_" & evt & "_deactivate", "DiverterEventHandler", 1000, Array("deactivate", Me, m_deactivate_events(evt))
        Next
        For Each evt in m_activation_switches
            AddPinEventListener evt & "_active", m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_activate_events.Keys()
            RemovePinEventListener m_activate_events(evt).EventName, m_name & "_" & evt & "_activate"
        Next
        For Each evt in m_deactivate_events.Keys()
            RemovePinEventListener m_deactivate_events(evt).EventName, m_name & "_" & evt & "_deactivate"
        Next
        For Each evt in m_activation_switches
            RemovePinEventListener evt & "_active", m_name & "_activate"
        Next
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
    End Sub

    Public Sub Activate()
        Log "Activating"
        m_active = True
        GetRef(m_action_cb)(1)
        If m_activation_time.Value > 0 Then
            SetDelay m_name & "_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_activation_time.Value
        End If
        DispatchPinEvent m_name & "_activating", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        m_active = False
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_deactivating", Null
    End Sub

    Public Sub BallSearch(phase)
        If m_exclude_from_ball_search = True Then
            Exit Sub
        End If
        Log "Ball Search, phase " & phase
        If m_active = False Then
            If Not IsEmpty(m_action_cb) Then
                m_active = True
                GetRef(m_action_cb)(1)
            End If
            SetDelay m_name & "ball_search_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_ball_search_hold_time.Value
        Else
            If Not IsEmpty(m_action_cb) Then
                m_active = False
                GetRef(m_action_cb)(0)
            End If
            SetDelay m_name & "ball_search_activate", "DiverterEventHandler", Array(Array("activate", Me), Null), m_ball_search_hold_time.Value
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DiverterEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim diverter : Set diverter = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set DiverterEventHandler = kwargs
                Else
                    DiverterEventHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
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
    If IsObject(args(1)) Then
        Set DiverterEventHandler = kwargs
    Else
        DiverterEventHandler = kwargs
    End If
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
    Private m_complete
    Private m_exclude_from_ball_search
    Private m_use_roth
    Private m_roth_array_index
    
    Private m_debug


    Public Property Get Name()
        Name = Replace(m_name, "drop_target_", "")
    End Property
	Public Property Let Switch(value)
		m_switch = value
		AddPinEventListener m_switch & "_active", m_name & "_switch_active", "DroptargetEventHandler", 1000, Array("switch_active", Me)
		AddPinEventListener m_switch & "_inactive", m_name & "_switch_inactive", "DroptargetEventHandler", 1000, Array("switch_inactive", Me)
	End Property
    Public Property Get Switch()
        Switch = m_switch
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
    Public Property Let ExcludeFromBallSearch(value) : m_exclude_from_ball_search = value : End Property
    Public Property Get UseRothDroptarget()
        UseRothDroptarget = m_use_roth
    End Property
    Public Property Let UseRothDroptarget(value)
        m_use_roth = value
    End Property
    Public Property Get RothDTSwitchID()
        RothDTSwitchID = m_roth_array_index
    End Property
    Public Property Let RothDTSwitchID(value)
        m_roth_array_index = value
    End Property
    
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "drop_target_" & name
		m_switch = Empty
        EnableKeepUpEvents = Array()
        DisableKeepUpEvents = Array()
		m_action_cb = Empty
		KnockdownEvents = Array()
		ResetEvents = Array()
        m_complete = 0
		m_debug = False
        m_use_roth = False
        m_roth_array_index = -1
        m_exclude_from_ball_search = False
        glf_drop_targets.Add name, Me
        Set Init = Me
	End Function

    Public Sub UpdateStateFromSwitch(is_complete)

		Log "Drop target " & m_name & " switch " & m_switch & " has active value " & is_complete & " compared to drop complete " & m_complete

		If is_complete <> m_complete Then
			If is_complete = 1 Then
				Down()
			Else
				Up()
			End	If
		End If
		'UpdateBanks()
    End Sub

    Public Sub Up()
        m_complete = 0
        DispatchPinEvent m_name & "_up", Null
    End Sub

	Public Sub Down()
        m_complete = 1
        DispatchPinEvent m_name & "_down", Null
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
        If Not IsEmpty(m_action_cb) And m_complete = 0 Then
            GetRef(m_action_cb)(1)
		End If
    End Sub

	Public Sub Reset()
        If Not IsEmpty(m_action_cb) And m_complete = 1 Then
            GetRef(m_action_cb)(0)
		End If
    End Sub

    Public Sub BallSearch(phase)
        If m_exclude_from_ball_search = True Then
            Exit Sub
        End If
        Log "Ball Search, phase " & phase
        If Not IsEmpty(m_action_cb) And m_complete = 0 Then
            GetRef(m_action_cb)(1) 'Knockdown
            SetDelay m_name & "ball_search_deactivate", "DroptargetEventHandler", Array(Array("reset", Me), Null), 100
		Else
            If Not IsEmpty(m_action_cb) And m_complete = 1 Then
                GetRef(m_action_cb)(0) 'Reset
                SetDelay m_name & "ball_search_deactivate", "DroptargetEventHandler", Array(Array("knockdown", Me), Null), 100
            End If
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
    Private m_switch
    Private m_action_cb
    Private m_debug

    Public Property Let Switch(value)
        m_switch = value
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
        m_switch = Empty
        m_debug = False
        glf_flippers.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        If Not IsEmpty(m_switch) Then
            AddPinEventListener m_switch & "_active", m_name & "_active", "FlipperEventHandler", 1000, Array("activate", Me)
            AddPinEventListener m_switch & "_inactive", m_name & "_inactive", "FlipperEventHandler", 1000, Array("deactivate", Me)
        End If
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Deactivate()
        Dim evt
        RemovePinEventListener m_switch & "_active", m_name & "_active"
        RemovePinEventListener m_switch & "_inactive", m_name & "_inactive"
    End Sub

    Public Sub Activate()
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(1)
        End If
        DispatchPinEvent m_name & "_activate", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
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
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
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
    If IsObject(args(1)) Then
        Set FlipperEventHandler = kwargs
    Else
        FlipperEventHandler = kwargs
    End If
End Function
Function CreateGlfLightSegmentDisplay(name)
	Dim segment_display : Set segment_display = (new GlfLightSegmentDisplay)(name)
	Set CreateGlfLightSegmentDisplay = segment_display
End Function

Class GlfLightSegmentDisplay
    private m_name
    private m_flash_on
    private m_flashing
    private m_flash_mask

    private m_text
    private m_current_text
    private m_display_state
    private m_current_state
    private m_current_flashing
    Private m_current_flash_mask
    private m_lights
    private m_light_group
    private m_light_groups
    private m_segmentmap
    private m_segment_type
    private m_size
    private m_update_method
    private m_text_stack
    private m_current_text_stack_entry
    private m_integrated_commas
    private m_integrated_dots
    private m_use_dots_for_commas
    private m_display_flash_duty
    private m_display_flash_display_flash_frequency
    private m_default_transition_update_hz
    private m_color
    private m_flex_dmd_index
    private m_b2s_dmd_index

    Public Property Get Name() : Name = m_name : End Property

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

    Public Property Get LightGroups() : LightGroups = m_light_groups : End Property
    Public Property Let LightGroups(input)
        m_light_groups = input
        CalculateLights()
    End Property

    Public Property Get UpdateMethod() : UpdateMethod = m_update_method : End Property
    Public Property Let UpdateMethod(input) : m_update_method = input : End Property

    Public Property Get SegmentSize() : SegmentSize = m_size : End Property
    Public Property Let SegmentSize(input)
        m_size = input
        CalculateLights()
    End Property

    Public Property Get IntegratedCommas() : IntegratedCommas = m_integrated_commas : End Property
    Public Property Let IntegratedCommas(input) : m_integrated_commas = input : End Property

    Public Property Get IntegratedDots() : IntegratedDots = m_integrated_dots : End Property
    Public Property Let IntegratedDots(input) : m_integrated_dots = input : End Property

    Public Property Get UseDotsForCommas() : UseDotsForCommas = m_use_dots_for_commas : End Property
    Public Property Let UseDotsForCommas(input) : m_use_dots_for_commas = input : End Property

    Public Property Get DefaultColor() : DefaultColor = m_color : End Property
    Public Property Let DefaultColor(input) : m_color = input : End Property

    Public Property Let ExternalFlexDmdSegmentIndex(input)
        m_flex_dmd_index = input
    End Property

    Public Property Let ExternalB2SSegmentIndex(input)
        m_b2s_dmd_index = input
    End Property

    Public Property Get DefaultTransitionUpdateHz() : DefaultTransitionUpdateHz = m_default_transition_update_hz : End Property
    Public Property Let DefaultTransitionUpdateHz(input) : m_default_transition_update_hz = input : End Property

    Public default Function init(name)
        m_name = name
        m_flash_on = True
        m_flashing = "no_flash"
        m_flash_mask = Empty
        m_text = Empty
        m_size = 0
        m_segment_type = Empty
        m_segmentmap = Null
        m_light_group = Empty
        m_light_groups = Array()
        m_current_text = Empty
        m_display_state = Empty
        m_current_state = Null
        m_current_flashing = Empty
        m_current_flash_mask = Empty
        m_current_text_stack_entry = Null
        Set m_text_stack = (new GlfTextStack)()
        m_update_method = "replace"
        m_lights = Array()  
        m_integrated_commas = False
        m_integrated_dots = False
        m_use_dots_for_commas = False
        m_flex_dmd_index = -1
        m_b2s_dmd_index = -1

        m_display_flash_duty = 30
        m_default_transition_update_hz = 30
        m_display_flash_display_flash_frequency = 60

        m_color = "ffffff"

        SetDelay m_name & "software_flash", "Glf_SegmentDisplaySoftwareFlashEventHandler", Array(True, Me), 100

        glf_segment_displays.Add name, Me
        Set Init = Me
    End Function

    Private Sub CalculateLights()
        If Not IsEmpty(m_segment_type) And m_size > 0 And (Not IsEmpty(m_light_group) Or Ubound(m_light_groups)>-1) Then
            m_lights = Array()
            If m_segment_type = "14Segment" Then
                ReDim m_lights((m_size * 15)-1)
            ElseIf m_segment_type = "7Segment" Then
                ReDim m_lights((m_size * 8)-1)
            End If

            Dim i, group_idx, current_light_group
            If Not IsEmpty(m_light_group) Then
                current_light_group = m_light_group
            ElseIf UBound(m_light_groups)>-1 Then
                current_light_group = m_light_groups(0)
                group_idx = 0
            End If
            Dim k : k = 0
            For i=0 to UBound(m_lights)
                'On Error Resume Next
                If typename(Eval(current_light_group & CStr(k+1))) = "Light" Then
                    m_lights(i) = current_light_group & CStr(k+1)
                    k=k+1
                Else
                    'msgbox typename(Eval(current_light_group & CStr(k+1)))
                    'msgbox current_light_group & CStr(k+1)
                    current_light_group = m_light_groups(group_idx+1)
                    'msgbox current_light_group
                    group_idx = group_idx + 1
                    k = 0
                    m_lights(i) = current_light_group & CStr(k+1)
                    k=k+1
                End If

            Next
        End If
    End Sub

    Public Sub SetVirtualDMDLights(input)
        If m_flex_dmd_index>-1 Then
            Dim x
            For x=0 to UBound(m_lights)
                glf_lightNames(m_lights(x)).Visible = input
            Next
        End If
    End Sub

    Private Sub SetText(text, flashing, flash_mask)
        'Set a text to the display.
        Exit Sub


        'If flashing = "no_flash" Then
        '    m_flash_on = True
        'ElseIf flashing = "flash_mask" Then
            'm_flash_mask = flash_mask.rjust(len(text))
        'End If

        'If flashing = "no_flash" or m_flash_on = True or Not IsNull(text) Then
        '    If text <> m_display_state Then
        '        m_display_state = text
                'Set text to lights.
        '        If text="" Then
        '            text = Glf_FormatValue(text, " >" & CStr(m_size))
        '        Else
        '            text = Right(text, m_size)
        '        End If
        '        If text <> m_current_text Then
        '            m_current_text = text
        '            UpdateText()
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub UpdateDisplay(segment_text, flashing, flash_mask)
        Set m_current_state = segment_text
        m_flashing = flashing
        m_flash_mask = flash_mask
        'SetText m_current_state.ConvertToString(), flashing, flash_mask
        UpdateText()
    End Sub

    Private Sub UpdateText()
        'iterate lights and chars
        Dim mapped_text, segment
        If m_flash_on = True Or m_flashing = "no_flash" Then
            mapped_text = MapSegmentTextToSegments(m_current_state, m_size, m_segmentmap)
        Else
            If m_flashing = "mask" Then
                mapped_text = MapSegmentTextToSegments(m_current_state.BlankSegments(m_flash_mask), m_size, m_segmentmap)
            ElseIf m_flashing = "match" Then
                mapped_text = MapSegmentTextToSegments(m_current_state.BlankSegments(String(m_size, "F")), m_size, m_segmentmap)
            Else
                mapped_text = MapSegmentTextToSegments(m_current_state.BlankSegments(String(m_size, "F")), m_size, m_segmentmap)
            End If
        End If
        Dim segment_idx, i : segment_idx = 0 : i = 0
        For Each segment in mapped_text
            
            If m_segment_type = "14Segment" Then
                Glf_SetLight m_lights(segment_idx), SegmentColor(segment.a)
                Glf_SetLight m_lights(segment_idx + 1), SegmentColor(segment.b)
                Glf_SetLight m_lights(segment_idx + 2), SegmentColor(segment.c)
                Glf_SetLight m_lights(segment_idx + 3), SegmentColor(segment.d)
                Glf_SetLight m_lights(segment_idx + 4), SegmentColor(segment.e)
                Glf_SetLight m_lights(segment_idx + 5), SegmentColor(segment.f)
                Glf_SetLight m_lights(segment_idx + 6), SegmentColor(segment.g1)
                Glf_SetLight m_lights(segment_idx + 7), SegmentColor(segment.g2)
                Glf_SetLight m_lights(segment_idx + 8), SegmentColor(segment.h)
                Glf_SetLight m_lights(segment_idx + 9), SegmentColor(segment.j)
                Glf_SetLight m_lights(segment_idx + 10), SegmentColor(segment.k)
                Glf_SetLight m_lights(segment_idx + 11), SegmentColor(segment.n)
                Glf_SetLight m_lights(segment_idx + 12), SegmentColor(segment.m)
                Glf_SetLight m_lights(segment_idx + 13), SegmentColor(segment.l)
                Glf_SetLight m_lights(segment_idx + 14), SegmentColor(segment.dp)
                If m_flex_dmd_index > -1 Then
                    'debug.print segment.CharMapping
                    dim hex
                    hex = segment.CharMapping
                    'debug.print typename(hex)
                    On Error Resume Next
				    glf_flex_alphadmd_segments(m_flex_dmd_index+i) = hex
				    If Err Then Debug.Print "Error: " & Err
                    'glf_flex_alphadmd_segments(m_flex_dmd_index+i) = segment.CharMapping '&h2A0F '0010101000001111
                    glf_flex_alphadmd.Segments = glf_flex_alphadmd_segments
                End If
                If m_b2s_dmd_index > -1 Then
                    dim b2sChar
                    b2sChar = segment.B2SLEDValue
                    On Error Resume Next
				    controller.B2SSetLED m_b2s_dmd_index+i, b2sChar
				    If Err Then Debug.Print "Error: " & Err
                End If
                segment_idx = segment_idx + 15
            ElseIf m_segment_type = "7Segment" Then
                Glf_SetLight m_lights(segment_idx), SegmentColor(segment.a)
                Glf_SetLight m_lights(segment_idx + 1), SegmentColor(segment.b)
                Glf_SetLight m_lights(segment_idx + 2), SegmentColor(segment.c)
                Glf_SetLight m_lights(segment_idx + 3), SegmentColor(segment.d)
                Glf_SetLight m_lights(segment_idx + 4), SegmentColor(segment.e)
                Glf_SetLight m_lights(segment_idx + 5), SegmentColor(segment.f)
                Glf_SetLight m_lights(segment_idx + 6), SegmentColor(segment.g)
                Glf_SetLight m_lights(segment_idx + 7), SegmentColor(segment.dp)
                segment_idx = segment_idx + 8
            End If
            i = i + 1
        Next
    End Sub

    Private Function SegmentColor(value)
        If value = 1 Then
            SegmentColor = m_color
        Else
            SegmentColor = "000000"
        End If
    End Function

    Public Sub AddTextEntry(text, color, flashing, flash_mask, transition, transition_out, priority, key)
    
        If m_update_method = "stack" Then
            m_text_stack.Push text,color,flashing,flash_mask,transition,transition_out,priority,key
            UpdateStack()
        Else
            Dim new_text : new_text = Glf_SegmentTextCreateCharacters(text.Value(), m_size, m_integrated_commas, m_integrated_dots, m_use_dots_for_commas, Array())
            Dim display_text : Set display_text = (new GlfSegmentDisplayText)(new_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, flashing, flash_mask
        End If
    End Sub

    Public Sub UpdateTransition(transition_runner)
        Dim display_text
        display_text = transition_runner.NextStep()
        If IsNull(display_text) Then
            UpdateStack() 
        Else
            Set display_text = (new GlfSegmentDisplayText)(display_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, m_flashing, m_flash_mask
            SetDelay m_name & "_update_transition", "Glf_SegmentDisplayUpdateTransition", Array(Me, transition_runner), 1000/m_default_transition_update_hz
        End If
    End Sub

    Public Sub UpdateStack()

        Dim top_text_stack_entry, top_is_current
        top_is_current = False
        If m_text_stack.IsEmpty() Then
            Dim empty_text : Set empty_text = (new GlfInput)("""" & String(m_size, " ") & """")
            Set top_text_stack_entry = (new GlfTextStackEntry)(empty_text,Null,"no_flash","",Null,Null,999999,"")
        Else
            Set top_text_stack_entry = m_text_stack.Peek()
        End If

        Dim previous_text_stack_entry : previous_text_stack_entry = Null
        If Not IsNull(m_current_text_stack_entry) Then
            Set previous_text_stack_entry = m_current_text_stack_entry
            If previous_text_stack_entry.text.IsPlayerState() Then
                RemovePlayerStateEventListener previous_text_stack_entry.text.PlayerStateValue(), m_name
            ElseIf previous_text_stack_entry.text.IsDeviceState() Then
                RemovePinEventListener top_text_stack_entry.text.DeviceStateEvent() , m_name
            End If

            If m_current_text_stack_entry.Key = top_text_stack_entry.Key Then
                top_is_current = True
            End If
        End If
        
        Set m_current_text_stack_entry = top_text_stack_entry

        'determine if the new key is different than the previous key (out transitions are only applied when changing keys)
        Dim transition_config : transition_config = Null
        If Not IsNull(previous_text_stack_entry) Then
            If top_text_stack_entry.key <> previous_text_stack_entry.key And Not IsNull(previous_text_stack_entry.transition_out) Then
                Set transition_config = previous_text_stack_entry.transition_out
            End If
        End If
        'determine if new text entry has a transition, if so, apply it (overrides any outgoing transition)
        If Not IsNull(top_text_stack_entry.transition) Then
            Set transition_config = top_text_stack_entry.transition
        End If
        'start transition (if configured)
        Dim flashing, flash_mask, display_text
        If Not IsNull(transition_config) And Not top_is_current Then
            'msgbox "starting transition"
            Dim transition_runner
            Select Case transition_config.TransitionType()
                case "push":
                    Set transition_runner = (new GlfPushTransition)(m_size, True, True, True)
                    transition_runner.Direction = transition_config.Direction()
                    transition_runner.TransitionText = transition_config.Text()
                case "cover":
                    Set transition_runner = (new GlfCoverTransition)(m_size, True, True, True)
                    transition_runner.Direction = transition_config.Direction()
                    transition_runner.TransitionText = transition_config.Text()
            End Select

            Dim previous_text
            If Not IsNull(previous_text_stack_entry) Then
                previous_text = previous_text_stack_entry.text.Value()
            Else
                previous_text = String(m_size, " ")
            End If

            If Not IsEmpty(top_text_stack_entry.flashing) Then
                flashing = top_text_stack_entry.flashing
                flash_mask = top_text_stack_entry.flash_mask
            Else
                flashing = m_current_state.flashing
                flash_mask = m_current_state.flash_mask
            End If
            display_text = transition_runner.StartTransition(previous_text, top_text_stack_entry.text.Value(), Array(), Array())
            Set display_text = (new GlfSegmentDisplayText)(display_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, flashing, flash_mask
            SetDelay m_name & "_update_transition", "Glf_SegmentDisplayUpdateTransition", Array(Me, transition_runner), 1000/m_default_transition_update_hz
        Else
            'no transition - subscribe to text template changes and update display
            RemoveDelay m_name & "_update_transition"
            If top_text_stack_entry.text.IsPlayerState() Then
                AddPlayerStateEventListener top_text_stack_entry.text.PlayerStateValue(), m_name, top_text_stack_entry.text.PlayerStatePlayer(), "Glf_SegmentTextStackEventHandler", top_text_stack_entry.priority, Me
            ElseIf top_text_stack_entry.text.IsDeviceState() Then
                AddPinEventListener top_text_stack_entry.text.DeviceStateEvent() , m_name, "Glf_SegmentTextStackEventHandler", top_text_stack_entry.priority, Me
            End If

            'set any flashing state specified in the entry
            If Not IsEmpty(top_text_stack_entry.flashing) Then
                flashing = top_text_stack_entry.flashing
                flash_mask = top_text_stack_entry.flash_mask
            Else
                flashing = m_current_state.flashing
                flash_mask = m_current_state.flash_mask
            End If

            'update the display
            Dim text_value : text_value = top_text_stack_entry.text.Value()

            If text_value = False Then
                text_value = String(m_size, " ")
            End If
            Dim new_text : new_text = Glf_SegmentTextCreateCharacters(text_value, m_size, m_integrated_commas, m_integrated_dots, m_use_dots_for_commas, Array())
            Set display_text = (new GlfSegmentDisplayText)(new_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, flashing, flash_mask
        End If
    End Sub

    Public Sub CurrentPlaceholderChanged()
        Dim text_value : text_value = m_current_text_stack_entry.text.Value()
        'msgbox text_value
        If text_value = False Then
            text_value = String(m_size, " ")
        End If
        Dim new_text : new_text = Glf_SegmentTextCreateCharacters(text_value, m_size, m_integrated_commas, m_integrated_dots, m_use_dots_for_commas, Array())
        Dim display_text : Set display_text = (new GlfSegmentDisplayText)(new_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
        UpdateDisplay display_text, m_current_text_stack_entry.flashing, m_current_text_stack_entry.flash_mask
    End Sub

    Public Sub RemoveTextByKey(key)
        m_text_stack.PopByKey key
        UpdateStack()
    End Sub

    Public Sub RemoveTextByKeyNoUpdate(key)
        m_text_stack.PopByKey key
    End Sub

    Public Sub SetFlashing(flash_type)

    End Sub

    Public Sub SetFlashingMask(mask)

    End Sub

    Public Sub SetColor(color)

    End Sub

    Public Sub SetSoftwareFlash(enabled)
        m_flash_on = enabled

        If m_flashing = "no_flash" Then
            Exit Sub
        End If

        If IsNull(m_current_state) Then
            Exit Sub
        End If
        UpdateText        
    End Sub

End Class

Sub Glf_SegmentDisplaySoftwareFlashEventHandler(args)
    Dim display, enabled
    Set display = args(1)
    enabled = args(0)
    If enabled = True Then
        SetDelay display.Name & "software_flash", "Glf_SegmentDisplaySoftwareFlashEventHandler", Array(False, display), 100
        display.SetSoftwareFlash True
    Else
        SetDelay display.Name & "software_flash", "Glf_SegmentDisplaySoftwareFlashEventHandler", Array(True, display), 100
        display.SetSoftwareFlash False
    End If

    
End Sub

Sub Glf_SegmentTextStackEventHandler(args)
    Dim segment
    Set segment = args(0) 
    'kwargs = args(1) 
    'Dim player_var : player_var = kwargs(0)
    'Dim value : value = kwargs(1)
    'Dim prevValue : prevValue = kwargs(2)
    segment.CurrentPlaceholderChanged()
End Sub

Sub Glf_SegmentDisplayUpdateTransition(args)
    Dim display, runner
    Set display = args(0) 
    Set runner = args(1)
    display.UpdateTransition runner
End Sub

Class GlfTextStackEntry
    Public text, colors, flashing, flash_mask, transition, transition_out, priority, key

    Public default Function init(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
        Set Me.text = text
        Me.colors = colors
        Me.flashing = flashing
        Me.flash_mask = flash_mask
        If Not IsNull(transition) Then
            Set Me.transition = transition
        Else
            Me.transition = Null
        End If
        If Not IsNull(transition_out) Then
            Set Me.transition_out = transition_out
        Else
            Me.transition_out = Null
        End If
        Me.priority = priority
        Me.key = key
        Set Init = Me
    End Function
End Class

Class GlfTextStack
    Private stack

    ' Initialize an empty stack
    Public default Function Init()
        ReDim stack(-1)  ' Initialize an empty array
        Set Init = Me
    End Function

    ' Push a new text entry onto the stack or update an existing one
    Public Sub Push(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
        Dim found : found = False
        Dim i

        ' Check if the key already exists in the stack and update it
        For i = LBound(stack) To UBound(stack)
            If stack(i).key = key Then
                ' Replace the existing item if the key matches
                Set stack(i) = CreateTextStackEntry(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            ' Insert the new item into the array maintaining priority order
            ReDim Preserve stack(UBound(stack) + 1)
            Set stack(UBound(stack)) = CreateTextStackEntry(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
            SortStackByPriority
        End If
    End Sub

    ' Pop a specific entry from the stack by key
    Public Function PopByKey(key)
        Dim i, removedItem, found
        found = False
        Set removedItem = Nothing
    
        ' Loop through the stack to find the item with the matching key
        For i = LBound(stack) To UBound(stack)
            If stack(i).key = key Then
                ' Store the item to be removed
                Set removedItem = stack(i)
                found = True
    
                ' Shift all elements after the removed item to the left
                Dim j
                For j = i To UBound(stack) - 1
                    Set stack(j) = stack(j + 1)
                Next
    
                ' Resize the array to remove the last element
                ReDim Preserve stack(UBound(stack) - 1)
                Exit For
            End If
        Next
    
        ' Return the removed item (or Nothing if not found)
        If found Then
            Set PopByKey = removedItem
        Else
            Set PopByKey = Nothing
        End If
    End Function

    ' Peek at the top entry of the stack without popping it
    Public Function Peek()
        If UBound(stack) >= 0 Then
            Set Peek = stack(LBound(stack))
        Else
            Set Peek = Nothing
        End If
    End Function

    ' Check if the stack is empty
    Public Function IsEmpty()
        IsEmpty = (UBound(stack) < 0)
    End Function

    ' Create a new GlfTextStackEntry object
    Private Function CreateTextStackEntry(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
        Dim entry
        Set entry = New GlfTextStackEntry
        entry.init text, colors, flashing, flash_mask, transition, transition_out, priority, key
        Set CreateTextStackEntry = entry
    End Function

    ' Sort the stack by priority (descending)
    Private Sub SortStackByPriority()
        Dim i, j
        Dim temp
        For i = LBound(stack) To UBound(stack) - 1
            For j = i + 1 To UBound(stack)
                If stack(i).priority < stack(j).priority Then
                    ' Swap the elements
                    Set temp = stack(i)
                    Set stack(i) = stack(j)
                    Set stack(j) = temp
                End If
            Next
        Next
    End Sub
End Class


Class FourteenSegments
    Public dp, l, m, n, k, j, h, g2, g1, f, e, d, c, b, a, char, hexcode, hexcode_dp

    Public Function CloneMapping()
        Set CloneMapping = (new FourteenSegments)(dp,l,m,n,k,j,h,g2,g1,f,e,d,c,b,a,char)
    End Function

    Public Property Get CharMapping()
        If dp = 1 Then
            CharMapping = hexcode_dp
        Else
            CharMapping = hexcode
        End If
    End Property

    Public Property Get B2SLEDValue()						'to be used with dB2S 15-segments-LED used in Herweh's Designer
        B2SLEDValue = 0									'default for unknown characters
        select case char
            Case "","":	B2SLEDValue = 0
            Case "0":	B2SLEDValue = 63	
            Case "1":	B2SLEDValue = 8704
            Case "2":	B2SLEDValue = 2139
            Case "3":	B2SLEDValue = 2127	
            Case "4":	B2SLEDValue = 2150
            Case "5":	B2SLEDValue = 2157
            Case "6":	B2SLEDValue = 2172
            Case "7":	B2SLEDValue = 7
            Case "8":	B2SLEDValue = 2175
            Case "9":	B2SLEDValue = 2159
            Case "A":	B2SLEDValue = 2167
            Case "B":	B2SLEDValue = 10767
            Case "C":	B2SLEDValue = 57
            Case "D":	B2SLEDValue = 8719
            Case "E":	B2SLEDValue = 121
            Case "F":	B2SLEDValue = 2161
            Case "G":	B2SLEDValue = 2109
            Case "H":	B2SLEDValue = 2166
            Case "I":	B2SLEDValue = 8713
            Case "J":	B2SLEDValue = 31
            Case "K":	B2SLEDValue = 5232
            Case "L":	B2SLEDValue = 56
            Case "M":	B2SLEDValue = 1334
            Case "N":	B2SLEDValue = 4406
            Case "O":	B2SLEDValue = 63
            Case "P":	B2SLEDValue = 2163
            Case "Q":	B2SLEDValue = 4287
            Case "R":	B2SLEDValue = 6259
            Case "S":	B2SLEDValue = 2157
            Case "T":	B2SLEDValue = 8705
            Case "U":	B2SLEDValue = 62
            Case "V":	B2SLEDValue = 17456
            Case "W":	B2SLEDValue = 20534
            Case "X":	B2SLEDValue = 21760
            Case "Y":	B2SLEDValue = 9472
            Case "Z":	B2SLEDValue = 17417
            Case "<":	B2SLEDValue = 5120
            Case ">":	B2SLEDValue = 16640
            Case "^":	B2SLEDValue = 17414
            Case ".":	B2SLEDValue = 8
            Case "!":	B2SLEDValue = 0
            Case ".":	B2SLEDValue = 128
            Case "*":	B2SLEDValue = 32576
            Case "/":	B2SLEDValue = 17408
            Case "\":	B2SLEDValue = 4352
            Case "|":	B2SLEDValue = 8704
            Case "=":	B2SLEDValue = 2120
            Case "+":	B2SLEDValue = 10816
            Case "-":	B2SLEDValue = 2112
        End Select			
        B2SLEDValue = cint(B2SLEDValue)
    End Property

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

        Dim binaryString, decimalValue, i
        binaryString = CStr("0" & n & m & l & g2 & k & j & h & dp & g1 & f & e & d & c & b & a)
        If binaryString = "0000000000000000" Then
            hexcode = 0
        Else
            decimalValue = 0
            For i = 1 To Len(binaryString)
                decimalValue = decimalValue * 2 + Mid(binaryString, i, 1)
            Next
            hexcode = CInt("&H" & Right("0000" & UCase(Hex(decimalValue)), 4))
        End If
        binaryString = CStr("0" & n & m & l & g2 & k & j & h & 1 & g1 & f & e & d & c & b & a)        
        decimalValue = 0
        For i = 1 To Len(binaryString)
            decimalValue = decimalValue * 2 + Mid(binaryString, i, 1)
        Next
        hexcode_dp = CInt("&H" & Right("0000" & UCase(Hex(decimalValue)), 4))
        Set Init = Me
    End Function
End Class



Class SevenSegments
    Public dp, g, f, e, d, c, b, a, char

    Public Function CloneMapping()
        Set CloneMapping = (new SevenSegments)(dp,g,f,e,d,c,b,a,char)
    End Function

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
FOURTEEN_SEGMENTS.Add 74, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, "J")
FOURTEEN_SEGMENTS.Add 75, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, "K")
FOURTEEN_SEGMENTS.Add 76, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, "L")
FOURTEEN_SEGMENTS.Add 77, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, "M")
FOURTEEN_SEGMENTS.Add 78, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, "N")
FOURTEEN_SEGMENTS.Add 79, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "O")
FOURTEEN_SEGMENTS.Add 80, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, 1, "P")
FOURTEEN_SEGMENTS.Add 81, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "Q")
FOURTEEN_SEGMENTS.Add 82, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, 1, "R")
FOURTEEN_SEGMENTS.Add 83, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 0, 1, "S")
FOURTEEN_SEGMENTS.Add 84, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, "T")
FOURTEEN_SEGMENTS.Add 85, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, "U")
FOURTEEN_SEGMENTS.Add 86, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, "V")
FOURTEEN_SEGMENTS.Add 87, (New FourteenSegments)(0, 1, 0, 1, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, "W")
FOURTEEN_SEGMENTS.Add 88, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "X")
FOURTEEN_SEGMENTS.Add 89, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1, 0, "Y")
FOURTEEN_SEGMENTS.Add 90, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, "Z")
FOURTEEN_SEGMENTS.Add 91, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, "[")
FOURTEEN_SEGMENTS.Add 92, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, Chr(92)) ' Character \
FOURTEEN_SEGMENTS.Add 93, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, "]")
FOURTEEN_SEGMENTS.Add 94, (New FourteenSegments)(0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "^")
FOURTEEN_SEGMENTS.Add 95, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, "_")
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


Function MapSegmentTextToSegments(text_state, display_width, segment_mapping)
    'Map a segment display text to a certain display mapping.

    Dim text : text = text_state.Text
    Dim segments()
    ReDim segments(UBound(text))

    Dim charCode, char, mapping, i, new_mapping
    For i = 0 To UBound(text)
        Set char = text(i)
        If segment_mapping.Exists(char("char_code")) Then
            Set mapping = segment_mapping(char("char_code"))
            Set new_mapping = mapping.CloneMapping()
            If char("dot") = True Then
                new_mapping.dp = 1
            Else
                new_mapping.dp = 0
            End If
        Else
            Set new_mapping = segment_mapping(Null)
        End If

        Set segments(i) = new_mapping
    Next

    MapSegmentTextToSegments = segments
End Function


Class GlfSegmentDisplayText

    Private m_embed_dots
    Private m_embed_commas
    Private m_use_dots_for_commas
    Private m_text

    Public Property Get Text() : Text = m_text : End Property

    ' Initialize the class
    Public default Function Init(char_list, embed_dots, embed_commas, use_dots_for_commas)
        m_embed_dots = embed_dots
        m_embed_commas = embed_commas
        m_use_dots_for_commas = use_dots_for_commas
        m_text = char_list
        Set Init = Me
    End Function

    ' Get the length of the text
    Public Function Length()
        Length = UBound(m_text) + 1
    End Function

    ' Get a character or a slice of the text
    Public Function GetItem(index)
        If IsArray(index) Then
            Dim slice, i
            slice = Array()
            For i = LBound(index) To UBound(index)
                slice = AppendArray(slice, m_text(index(i)))
            Next
            GetItem = slice
        Else
            GetItem = m_text(index)
        End If
    End Function

    ' Extend the text with another list
    Public Sub Extend(other_text)
        Dim i
        For i = LBound(other_text) To UBound(other_text)
            m_text = AppendArray(m_text, other_text(i))
        Next
    End Sub

    ' Convert the text to a string
    Public Function ConvertToString()
        Dim text, char, i
        text = ""
        For i = LBound(m_text) To UBound(m_text)
            Set char = m_text(i)
            text = text & Chr(char("char_code"))
            If char("dot") Then text = text & "."
            If char("comma") Then text = text & ","
        Next
        ConvertToString = text
    End Function

    ' Get colors (to be implemented in subclasses)
    Public Function GetColors()
        GetColors = Null
    End Function

    Public Function BlankSegments(flash_mask)
        Dim arrFlashMask, i
        ReDim arrFlashMask(Len(flash_mask) - 1)
        For i = 1 To Len(flash_mask)
            arrFlashMask(i - 1) = Mid(flash_mask, i, 1)
        Next


        Dim new_text, char, mask
        new_text = Array()
    
        ' Iterate over characters and the flash mask
        For i = LBound(m_text) To UBound(m_text)
            Set char = m_text(i)
            mask = arrFlashMask(i)
    
            ' If mask is "F", blank the character
            If mask = "F" Then
                new_text = AppendArray(new_text, Glf_SegmentTextCreateDisplayCharacter(32, False, False, char("color")))
            Else
                ' Otherwise, keep the character as is
                new_text = AppendArray(new_text, char)
            End If
        Next
    
        ' Create a new GlfSegmentDisplayText object with the updated text
        Dim blanked_text
        Set blanked_text = (new GlfSegmentDisplayText)(new_text, m_embed_commas, m_embed_dots, m_use_dots_for_commas)
        Set BlankSegments = blanked_text
    End Function
    

End Class

' Helper function to append to an array
Function AppendArray(arr, value)
    Dim newArr, i
    ReDim newArr(UBound(arr) + 1)
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            Set newArr(i) = arr(i)
        Else
            newArr(i) = arr(i)
        End If
    Next
    If IsObject(value) Then
        Set newArr(UBound(newArr)) = value
    Else
        newArr(UBound(newArr)) = value
    End If
    AppendArray = newArr
End Function

' Helper function to slice an array
Function SliceArray(arr, start_idx, end_idx)
    Dim sliced, i, j
    ReDim sliced(end_idx - start_idx)
    j = 0
    For i = start_idx To end_idx
        If IsObject(arr(i)) Then
            Set sliced(j) = arr(i)
        Else
            sliced(j) = arr(i)
        End If
        j = j + 1
    Next
    SliceArray = sliced
End Function

' Helper function to prepend an element to an array
Function PrependArray(arr, value)
    Dim newArr, i
    ReDim newArr(UBound(arr) + 1)
    If IsObject(value) Then
        Set newArr(0) = value
    Else
        newArr(0) = value
    End If
    
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            Set newArr(i + 1) = arr(i)
        Else
            newArr(i + 1) = arr(i)
        End If
    Next
    PrependArray = newArr
End Function


Function Glf_SegmentTextCreateCharacters(text, display_size, collapse_dots, collapse_commas, use_dots_for_commas, colors)
            


    Dim char_list, uncolored_chars, left_pad_color, default_right_color, i, char, color, current_length
    char_list = Array()

    ' Determine padding and default colors
    If IsArray(colors) And UBound(colors) >= 0 Then

        left_pad_color = colors(0)
        default_right_color = colors(UBound(colors))

    Else
        left_pad_color = Null
        default_right_color = Null
    End If

    ' Embed dots and commas
    uncolored_chars = Glf_SegmentTextEmbedDotsAndCommas(text, collapse_dots, collapse_commas, use_dots_for_commas)
 
    ' Adjust colors to match the uncolored characters
    If IsArray(colors) And UBound(colors) >= 0 Then
        Dim adjusted_colors
        adjusted_colors = SliceArray(colors, UBound(colors) - UBound(uncolored_chars) + 1, UBound(colors))
    Else
        adjusted_colors = Array()
    End If

    ' Create display characters
    For i = LBound(uncolored_chars) To UBound(uncolored_chars)
        char = uncolored_chars(i)
        
        If IsArray(adjusted_colors) And UBound(adjusted_colors) >= 0 Then
            color = adjusted_colors(0)
            adjusted_colors = SliceArray(adjusted_colors, 1, UBound(adjusted_colors))
        Else
            color = default_right_color
        End If
        char_list = AppendArray(char_list, Glf_SegmentTextCreateDisplayCharacter(char(0), char(1), char(2), color))
    Next

    ' Adjust the list size to match the display size
    current_length = UBound(char_list) + 1
    
    If current_length > display_size Then
        ' Truncate characters from the left
        char_list = SliceArray(char_list, current_length - display_size, UBound(char_list))
    ElseIf current_length < display_size Then
        ' Pad with spaces to the left
        Dim padding
        padding = display_size - current_length
        For i = 1 To padding
            char_list = PrependArray(char_list, Glf_SegmentTextCreateDisplayCharacter(32, False, False, left_pad_color))
        Next
    End If
    'msgbox ">"&text&"<"
    'msgbox UBound(char_list)
    Glf_SegmentTextCreateCharacters = char_list
End Function

Function Glf_SegmentTextEmbedDotsAndCommas(text, collapse_dots, collapse_commas, use_dots_for_commas)
    Dim char_has_dot, char_has_comma, char_list
    Dim i, char_code

    char_has_dot = False
    char_has_comma = False
    char_list = Array()

    ' Iterate through the text in reverse
    For i = Len(text) To 1 Step -1
        char_code = Asc(Mid(text, i, 1))
        
        ' Check for dots and commas and handle collapsing rules
        If (collapse_dots And char_code = Asc(".")) Or (use_dots_for_commas And char_code = Asc(",")) Then
            char_has_dot = True
        ElseIf collapse_commas And char_code = Asc(",") Then
            char_has_comma = True
        Else
            ' Insert the character at the start of the list
            char_list = PrependArray(char_list, Array(char_code, char_has_dot, char_has_comma))
            char_has_dot = False
            char_has_comma = False
        End If
    Next

    Glf_SegmentTextEmbedDotsAndCommas = char_list
End Function

' Helper function to create a display character
Function Glf_SegmentTextCreateDisplayCharacter(char_code, has_dot, has_comma, color)
    Dim display_character
    Set display_character = CreateObject("Scripting.Dictionary")
    display_character.Add "char_code", char_code
    display_character.Add "dot", has_dot
    display_character.Add "comma", has_comma
    display_character.Add "color", color
    Set Glf_SegmentTextCreateDisplayCharacter = display_character
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
        DispatchPinEvent m_name & "_flinging_ball", Null
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
        DispatchPinEvent m_name & "_releasing_ball", Null
        GetRef(m_action_cb)(0)
        SetDelay m_name & "_release_done", "MagnetEventHandler" , Array(Array("release_done", Me),Null), m_release_time.Value
    End Sub

    Public Sub ReleaseDone()
        m_release_in_progress = False
        DispatchPinEvent m_name & "_released_ball", Null
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
Class GlfCoverTransition

    Private m_output_length
    Private m_collapse_dots
    Private m_collapse_commas
    Private m_use_dots_for_commas

    Private m_current_step
    Private m_total_steps
    Private m_direction
    Private m_transition_text
    Private m_transition_color

    Private m_current_text
    Private m_new_text
    Private m_current_colors
    Private m_new_colors

    ' Initialize the transition class
    Public Default Function Init(output_length, collapse_dots, collapse_commas, use_dots_for_commas)
        m_output_length = output_length
        m_collapse_dots = collapse_dots
        m_collapse_commas = collapse_commas
        m_use_dots_for_commas = use_dots_for_commas

        m_current_step = 0
        m_total_steps = 0
        m_direction = "right"
        m_transition_text = "" ' Default empty transition text
        m_transition_color = Empty

        Set Init = Me
    End Function

    ' Set transition direction
    Public Property Let Direction(value)
        If value = "left" Or value = "right" Then
            m_direction = value
        Else
            m_direction = "right" ' Default to right
        End If
    End Property

    ' Set transition text
    Public Property Let TransitionText(value)
        m_transition_text = value
    End Property

    ' Start transition
    Public Function StartTransition(current_text, new_text, current_colors, new_colors)
        ' Store text and colors
        m_current_text = current_text
        m_new_text = Space(m_output_length - Len(new_text)) & new_text
        m_current_colors = current_colors
        m_new_colors = new_colors
        ' Calculate total steps for transition
        m_total_steps = m_output_length + Len(m_transition_text)
        'If m_total_steps > 0 Then m_total_steps = m_total_steps
        ' Reset step counter
        m_current_step = 0
        StartTransition = NextStep()
    End Function

    ' Manually call this to progress the transition
    Public Function NextStep()
        If m_current_step >= m_total_steps Then
            NextStep = Null
            Exit Function
        End If
        ' Get the correct transition text for this step
        Dim result
        result = GetTransitionStep(m_current_step)
        ' Increment step counter
        m_current_step = m_current_step + 1

        ' Return the current frame's text
        NextStep = result
    End Function

    ' Get the text for the current step
    Private Function GetTransitionStep(current_step)
        Dim transition_sequence, start_idx, end_idx

        If m_direction = "right" Then
            ' Right cover transition: new_text + transition_text moves in
            transition_sequence = m_new_text & m_transition_text & m_current_text
            start_idx = Len(transition_sequence) - (current_step + m_output_length)
            end_idx = start_idx + m_output_length - 1
        ElseIf m_direction = "left" Then
            ' Left cover transition: transition_text + new_text moves in
            transition_sequence = m_current_text & m_transition_text & m_new_text
            start_idx = current_step
            end_idx = start_idx + m_output_length - 1
        End If

        ' Ensure valid slice indices
        If start_idx < 0 Then start_idx = 0
        If end_idx > Len(transition_sequence) - 1 Then end_idx = Len(transition_sequence) - 1

        ' Extract the correct frame of text
        Dim sliced_text
        sliced_text = Mid(transition_sequence, start_idx + 1, end_idx - start_idx + 1)
        If m_output_length>m_current_step Then
            If m_direction = "right" Then
                sliced_text = m_new_text & m_transition_text & Right(m_current_text, m_output_length-current_step)
            ElseIf m_direction = "left" Then
                sliced_text = Left(m_current_text, m_output_length - m_current_step) & Left(m_transition_text & m_new_text, current_step)
            End If
        End If
        'MsgBox "transition_text-"&transition_sequence&", current_step=" & current_step & ", start_idx=" & start_idx & ", end_idx=" & end_idx & ", text=>" & sliced_text &"<, text_len=>" & Len(sliced_text) &"<, Total Steps: "&m_total_steps
        
        ' Convert only the final sliced text to segment display characters
        GetTransitionStep = Glf_SegmentTextCreateCharacters(sliced_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, m_new_colors)
    End Function

End Class
Class GlfPushTransition

    Private m_output_length
    Private m_collapse_dots
    Private m_collapse_commas
    Private m_use_dots_for_commas

    Private m_current_step
    Private m_total_steps
    Private m_direction
    Private m_transition_text
    Private m_transition_color

    Private m_current_text
    Private m_new_text
    Private m_current_colors
    Private m_new_colors

    ' Initialize the transition class
    Public Default Function Init(output_length, collapse_dots, collapse_commas, use_dots_for_commas)
        m_output_length = output_length
        m_collapse_dots = collapse_dots
        m_collapse_commas = collapse_commas
        m_use_dots_for_commas = use_dots_for_commas

        m_current_step = 0
        m_total_steps = 0
        m_direction = "right"
        m_transition_text = "" ' Default empty transition text
        m_transition_color = Empty

        Set Init = Me
    End Function

    ' Set transition direction
    Public Property Let Direction(value)
        If value = "left" Or value = "right" Then
            m_direction = value
        Else
            m_direction = "right" ' Default to right
        End If
    End Property

    ' Set transition text
    Public Property Let TransitionText(value)
        m_transition_text = value
    End Property

    ' Start transition
    Public Function StartTransition(current_text, new_text, current_colors, new_colors)
        ' Store text and colors
        m_current_text = current_text
        m_new_text = Space(m_output_length - Len(new_text)) & new_text
        m_current_colors = current_colors
        m_new_colors = new_colors
        ' Calculate total steps for transition
        m_total_steps = m_output_length + Len(m_transition_text)
        If m_total_steps > 0 Then m_total_steps = m_total_steps + 1
        'm_total_steps=(m_output_length*2)+1
        ' Reset step counter
        m_current_step = 0
        StartTransition = NextStep()
    End Function

    ' Manually call this to progress the transition
    Public Function NextStep()
        If m_current_step >= m_total_steps Then
            NextStep = Null ' Transition complete
            Exit Function
        End If
        ' Get the correct transition text for this step
        Dim result
        result = GetTransitionStep(m_current_step)
        ' Increment step counter
        m_current_step = m_current_step + 1

        ' Return the current frame's text
        NextStep = result
    End Function

    ' Get the text for the current step
    Private Function GetTransitionStep(current_step)
        Dim transition_sequence, start_idx, end_idx
        'msgbox "Step"
        ' Construct the full transition sequence as plain text
        If m_direction = "right" Then
            ' Right push: [NEW_TEXT + TRANSITION_TEXT + OLD_TEXT] moves LEFT
            transition_sequence = m_new_text & m_transition_text & m_current_text
    
            ' Calculate slice indices
            start_idx = Len(transition_sequence) - (current_step + m_output_length)
            end_idx = start_idx + m_output_length - 1
    
        ElseIf m_direction = "left" Then
            ' Left push: [OLD_TEXT + TRANSITION_TEXT + NEW_TEXT] moves RIGHT
            transition_sequence = m_current_text & m_transition_text & m_new_text
    
            ' Calculate slice indices
            start_idx = current_step
            end_idx = start_idx + m_output_length - 1
        End If
    
        ' Ensure valid slice indices
        If start_idx < 0 Then start_idx = 0
        If end_idx > Len(transition_sequence) - 1 Then end_idx = Len(transition_sequence) - 1
    
        ' Extract the correct frame of text
        Dim sliced_text
        sliced_text = Mid(transition_sequence, start_idx + 1, end_idx - start_idx + 1)
    
        ' Debugging output
        'MsgBox "transition_text-"&transition_sequence&", current_step=" & current_step & ", start_idx=" & start_idx & ", end_idx=" & end_idx & ", text=>" & sliced_text &"<"
        ' Convert only the final sliced text to segment display characters
        GetTransitionStep = Glf_SegmentTextCreateCharacters(sliced_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, m_new_colors)
    End Function

End Class


Function CreateGlfSoundBus(name)
	Dim bus : Set bus = (new GlfSoundBus)(name)
	Set CreateGlfSoundBus = bus
End Function

Class GlfSoundBus

    Private m_name
    Private m_simultaneous_sounds
    Private m_current_sounds
    Private m_volume
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "simultaneous_sounds":
                GetValue = m_simultaneous_sounds
            Case "volume":
                GetValue = m_volume
        End Select
    End Property

    Public Property Get SimultaneousSounds(): SimultaneousSounds = m_simultaneous_sounds: End Property
    Public Property Let SimultaneousSounds(input): m_simultaneous_sounds = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Let Debug(value)
        m_debug = value
    End Property

    Public default Function Init(name)
        m_name = "sound_bus_" & name
        m_simultaneous_sounds = 8
        m_volume = 0.5
        Set m_current_sounds = CreateObject("Scripting.Dictionary")
        glf_sound_buses.Add name, Me
        Set Init = Me
    End Function

    Public Sub Play(sound_settings)
        If (UBound(m_current_sounds.Keys)-1) > m_simultaneous_sounds Then
            'TODO: Queue Sound
        Else
            If m_current_sounds.Exists(sound_settings.Sound.File) Then
                m_current_sounds.Remove sound_settings.Sound.File
                RemoveDelay m_name & "_stop_sound_" & sound_settings.Sound.File       
            End If
            m_current_sounds.Add sound_settings.Sound.File, sound_settings
            Dim volume : volume = m_volume
            If Not IsEmpty(sound_settings.Sound.Volume) Then
                volume = sound_settings.Sound.Volume
            End If
            If Not IsEmpty(sound_settings.Volume) Then
                volume = sound_settings.Volume
            End If
            Dim loops : loops = 0
            If Not IsEmpty(sound_settings.Sound.Loops) Then
                loops = sound_settings.Sound.Loops
            End If
            If Not IsEmpty(sound_settings.Loops) Then
                loops = sound_settings.Loops
            End If

            PlaySound sound_settings.Sound.File, loops, volume, 0,0,0,0,0,0
            If loops = 0 Then
                SetDelay m_name & "_stop_sound_" & sound_settings.Sound.File, "Glf_SoundBusStopSoundHandler", Array(sound_settings.Sound.File, Me), sound_settings.Sound.Duration
            ElseIf loops>0 Then
                SetDelay m_name & "_stop_sound_" & sound_settings.Sound.File, "Glf_SoundBusStopSoundHandler", Array(sound_settings.Sound.File, Me), sound_settings.Sound.Duration*loops
            End If
        End If
    End Sub

    Public Sub StopSoundWithKey(sound_key)
        If m_current_sounds.Exists(sound_key) Then
            Dim sound_settings : Set sound_settings = m_current_sounds(sound_key)
            StopSound(sound_key)
            Dim evt
            For Each evt in sound_settings.Sound.EventsWhenStopped.Items()
                If evt.Evaluate() Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
            m_current_sounds.Remove sound_key
        End If
    End Sub

    Private Sub Log(message)
        If m_debug Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function Glf_SoundBusStopSoundHandler(args)
    Dim sound_key : sound_key = args(0)
    Dim sound_bus : Set sound_bus = args(1)
    sound_bus.StopSoundWithKey sound_key
End Function

Function CreateGlfSound(name)
	Dim sound : Set sound = (new GlfSound)(name)
	Set CreateGlfSound = sound
End Function

Class GlfSound

    Private m_name
    Private m_file
    Private m_events_when_stopped
    Private m_bus
    Private m_volume
    Private m_priority
    Private m_max_queue_time
    Private m_duration
    Private m_loops
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "file":
                GetValue = m_file
            Case "volume":
                GetValue = m_volume
            Case "events_when_stopped":
                Set GetValue = m_events_when_stopped
            Case "bus":
                GetValue = m_bus
            Case "priority":
                GetValue = m_priority
            Case "max_queue_time":
                GetValue = m_max_queue_time
            Case "duration":
                GetValue = m_duration
        End Select
    End Property

    Public Property Get File(): File = m_file: End Property
    Public Property Let File(input): m_file = input: End Property

    Public Property Get Bus(): Bus = m_bus: End Property
    Public Property Let Bus(input): m_bus = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input): m_priority = input: End Property

    Public Property Get MaxQueueTime(): MaxQueueTime = m_max_queue_time: End Property
    Public Property Let MaxQueueTime(input): m_max_queue_time = input: End Property

    Public Property Get Duration(): Duration = m_duration: End Property
    Public Property Let Duration(input): m_duration = input: End Property

    Public Property Get EventsWhenStopped(): Set EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_stopped.Add x, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
    End Property

    Public default Function Init(name)
        m_name = "sound_bus_" & name
        m_file = Empty
        m_bus = Empty
        m_volume = Empty
        m_priority = 0
        m_duration = 0
        m_max_queue_time = -1 
        m_loops = 0
        Set m_events_when_stopped = CreateObject("Scripting.Dictionary")
        glf_sounds.Add name, Me
        Set Init = Me
    End Function

    Private Sub Log(message)
        If m_debug Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class
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

Class GlfEvent
	Private m_raw, m_name, m_event, m_condition, m_delay, m_priority, has_condition
  
    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventName() : EventName = m_event : End Property
    Public Property Get Condition() : Condition = m_condition : End Property
    Public Property Get Delay() : Delay = m_delay : End Property
    Public Property Get Raw() : Raw = m_raw : End Property
    Public Property Get Priority() : Priority = m_priority : End Property

    Public Function Evaluate()
        If has_condition = True Then
            Evaluate = GetRef(m_condition)(Null)
        Else
            Evaluate = True
        End If
    End Function

	Public default Function init(evt)
        m_raw = evt
        Dim parsedEvent : parsedEvent = Glf_ParseEventInput(evt)
        m_name = parsedEvent(0)
        m_event = parsedEvent(1)
        m_condition = parsedEvent(2)
        If Not IsNull(m_condition) Then
            has_condition = True
        Else
            has_condition = False
        End If
        m_delay = parsedEvent(3)
        m_priority = parsedEvent(4)
	    Set Init = Me
	End Function

End Class

Class GlfEventDispatch
	Private m_raw, m_event, m_kwargs_ref
  
    Public Property Get EventName() : EventName = m_event : End Property
    Public Property Get Kwargs()
        If IsEmpty(m_kwargs_ref) Then
            Kwargs = Null
        Else
            Set Kwargs = GetRef(m_kwargs_ref)(Null)
        End If
    End Property
    Public Property Get Raw() : Raw = m_raw : End Property

	Public default Function init(evt)
        m_raw = evt
        Dim parsedDispatchEvent : parsedDispatchEvent = Glf_ParseDispatchEventInput(evt)
        m_event = parsedDispatchEvent(0)
        m_kwargs_ref = parsedDispatchEvent(1)
	    Set Init = Me
	End Function

End Class

Class GlfRandomEvent
	
    Private m_parent_key
    Private m_key
    Private m_mode
    Private m_events
    Private m_weights
    Private m_eventIndexMap
    Private m_fallback_event
    Private m_force_all
    Private m_force_different
    Private m_disable_random
    Private m_total_weights

    Public Property Let FallbackEvent(value) : m_fallback_event = value : End Property
    Public Property Let ForceAll(value) : m_force_all = value : End Property
    Public Property Let ForceDifferent(value) : m_force_different = value : End Property
    Public Property Let DisableRandom(value) : m_disable_random = value : End Property

	Public default Function init(evt, mode, key)
        m_parent_key = evt
        m_key = key
        m_mode = mode
        m_fallback_event = Empty
        m_force_all = True
        m_force_different = True
        m_disable_random = False
        m_total_weights = 0
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_weights = CreateObject("Scripting.Dictionary")
        Set m_eventIndexMap = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

    Public Sub Add(evt, weight)
        Dim newEvent : Set newEvent = (new GlfEvent)(evt)
        m_events.Add newEvent.Raw, newEvent
        m_weights.Add newEvent.Raw, weight
        m_total_weights = m_total_weights + weight
        m_eventIndexMap.Add newEvent.Raw, UBound(m_events.Keys)
    End Sub

    Public Function GetNextRandomEvent()

        Dim valid_events, event_to_fire
        Dim event_keys, event_items
        Dim i, count, key
        
        Set valid_events = CreateObject("Scripting.Dictionary")
        event_keys = m_events.Keys
        For i = 0 To UBound(event_keys)
            If m_events(event_keys(i)).Evaluate Then
                valid_events.Add event_keys(i), m_events(event_keys(i))
            End If
        Next
        event_to_fire = CheckFallback(valid_events)
        If Not IsEmpty(event_to_fire) Then
            GetNextRandomEvent = event_to_fire
            Exit Function
        End If

        If m_force_all = True Then
            event_keys = valid_events.Keys
            event_items = valid_events.Items
            valid_events.RemoveAll
            For i=0 to UBound(event_keys)
                If GetPlayerState("random_" & m_mode & "_" & m_key & "_" & m_eventIndexMap(event_keys(i))) = False Then
                    valid_events.Add event_keys(i), event_items(i)
                End If
            Next
        End If

        event_to_fire = CheckFallback(valid_events)
        If Not IsEmpty(event_to_fire) Then
            GetNextRandomEvent = event_to_fire
            Exit Function
        End If

        If m_force_different = True Then
            If valid_events.Exists(GetPlayerState("random_" & m_mode & "_" & m_key & "_last")) Then
                valid_events.Remove GetPlayerState("random_" & m_mode & "_" & m_key & "_last")
            End If
        End If

        event_to_fire = CheckFallback(valid_events)
        If Not IsEmpty(event_to_fire) Then
            GetNextRandomEvent = event_to_fire
            Exit Function
        End If

        If UBound(valid_events.Keys) = -1 Then
            GetNextRandomEvent = Empty
            Exit Function
        End If

        'Random Selection From remaining valid events
        Dim chosenKey
        If m_disable_random = False Then
            Dim total_weight
            For Each key In valid_events.Keys
                total_weight = total_weight + m_weights(key)
            Next

            Randomize
            'randomIdx = Int(Rnd() * (UBound(valid_events.Keys)-LBound(valid_events.Keys) + 1) + LBound(valid_events.Keys))
            Dim randVal
            randVal = Rnd() * total_weight
            Dim cumulativeWeight
            cumulativeWeight = 0
            
            For Each key In valid_events.Keys
                cumulativeWeight = cumulativeWeight + m_weights(key)
                If randVal <= cumulativeWeight Then
                    chosenKey = key
                    Exit For
                End If
            Next


        Else
            event_keys = m_events.Keys
            count = 0
            For i = 0 To UBound(event_keys)
                If GetPlayerState("random_" & m_mode & "_" & m_key & "_" & m_events(event_keys(i)).Raw) = True Then
                    If valid_events.Exists(m_events(event_keys(i)).Raw) Then
                        valid_events.Remove m_events(event_keys(i)).Raw
                    End If
                End If
            Next
            Dim valid_event_keys : valid_event_keys = valid_events.keys()
            chosenKey = valid_event_keys(0)
        End If
        
        SetPlayerState "random_" & m_mode & "_" & m_key & "_last", valid_events(chosenKey).Raw
        SetPlayerState "random_" & m_mode & "_" & m_key & "_" & valid_events(chosenKey).Raw, True

        event_keys = m_events.Keys
        count = 0
        For i = 0 To UBound(event_keys)
            If GetPlayerState("random_" & m_mode & "_" & m_key & "_" & m_events(event_keys(i)).Raw) = True Then
                count = count + 1
            End If
        Next
        If count = (UBound(event_keys) + 1) Then
            For i = 0 To UBound(event_keys)
                SetPlayerState "random_" & m_mode & "_" & m_key & "_" & m_events(event_keys(i)).Raw, False
            Next
        End If

        GetNextRandomEvent = valid_events(chosenKey).EventName

    End Function

    Public Function CheckFallback(valid_events)
        If UBound(valid_events.Keys()) = -1 Then
            If Not IsEmpty(m_fallback_event) Then
                CheckFallback = m_fallback_event
            Else
                CheckFallback = Empty
            End If
        Else
            CheckFallback = Empty
        End If
    End Function

    Public Function ToYaml
        Dim yaml
        yaml = yaml & "    events:" & vbCrLf
        For Each evt in m_events.Keys
            yaml = yaml & "      " & evt & ": " & m_weights(evt) & vbCrLf
        Next
        ToYaml = yaml

    End Function

End Class


'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub Glf_AddPlayer()
    Dim kwargs : Set kwargs = GlfKwargs()
    With kwargs
        .Add "num", -1
    End With
    Select Case UBound(glf_playerState.Keys())
        Case -1:
            kwargs("num") = 1
            DispatchPinEvent "player_added", kwargs
            glf_playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            SetPlayerStateByPlayer GLF_SCORE, 0, 0
            SetPlayerStateByPlayer "number", 1, 0
            Glf_BcpAddPlayer 1
            glf_currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 2
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 2", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 1
                SetPlayerStateByPlayer "number", 2, 1
                Glf_BcpAddPlayer 2
            End If
        Case 1:
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 3
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 3", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 2
                SetPlayerStateByPlayer "number", 3, 2
                Glf_BcpAddPlayer 3
            End If     
        Case 2:   
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 4
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 4", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 3
                SetPlayerStateByPlayer "number", 4, 3
                Glf_BcpAddPlayer 4
            End If  
            glf_canAddPlayers = False
    End Select
End Sub

Function Glf_InitNewPlayer()
    Dim state : Set state = CreateObject("Scripting.Dictionary")
    state.Add GLF_SCORE, 0
    Glf_MonitorPlayerStateUpdate GLF_SCORE, 0
    state.Add GLF_INITIALS, ""
    Glf_MonitorPlayerStateUpdate GLF_INITIALS, ""
    state.Add GLF_CURRENT_BALL, 1
    Glf_MonitorPlayerStateUpdate GLF_CURRENT_BALL, 1
    state.Add "extra_balls", 0
    Glf_MonitorPlayerStateUpdate "extra_balls", 0
    Dim i
    Dim init_var_keys : init_var_keys = glf_initialVars.Keys()
    Dim init_var_items : init_var_items = glf_initialVars.Items()
    For i=0 To UBound(init_var_keys)
        state.Add init_var_keys(i), init_var_items(i)
        Glf_MonitorPlayerStateUpdate init_var_keys(i), init_var_items(i)
    Next
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

AddPinEventListener "request_to_start_game", "request_to_start_game_ball_controller", "Glf_BallController", 30, Null
Function Glf_BallController(args)
    Dim balls_in_trough : balls_in_trough = 0
    If glf_troughSize = 1 Then
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 2 Then
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 3 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 4 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 5 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 6 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 7 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough7.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 8 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough7.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If Drain.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
  
    If glf_troughSize <> balls_in_trough Then
        Glf_BallController = False
        Exit Function
    End If

    Glf_BallController = True
End Function

AddPinEventListener "request_to_start_game", "request_to_start_game_result", "Glf_StartGame", 20, Null

Function Glf_StartGame(args)
    If args(1) = True And glf_gameStarted = False Then
        Glf_AddPlayer()
        glf_gameStarted = True
        glf_canAddPlayers = True
        DispatchPinEvent GLF_GAME_START, Null
        SetDelay GLF_GAME_STARTED, "Glf_DispatchGameStarted", Null, 50
    End If
End Function

Sub Glf_EndBall()

    glf_BIP = 0
    DispatchPinEvent GLF_BALL_WILL_END, Null
    DispatchQueuePinEvent GLF_BALL_ENDING, Null
    Dim device
    For Each device in glf_flippers.Items()
        device.Disable()
    Next
    For Each device in glf_autofiredevices.Items()
        device.Disable()
    Next

End Sub

Public Function Glf_DispatchGameStarted(args)
    DispatchPinEvent GLF_GAME_STARTED, Null
End Function


'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener GLF_GAME_STARTED, "start_game_release_ball",   "Glf_ReleaseBall", 20, True
AddPinEventListener GLF_NEXT_PLAYER, "next_player_release_ball",   "Glf_ReleaseBall", 20, True
'
'*****************************
Function Glf_ReleaseBall(args)
    Dim kwargs
    Set kwargs = GlfKwargs()
    If Not IsNull(args) Then
        If args(0) = True Then
            kwargs.Add "new_ball", True
        End If
    End If
    DispatchQueuePinEvent "balldevice_trough_ball_eject_attempt", kwargs
End Function


'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener "balldevice_trough_ball_eject_attempt", "trough_eject",  "Glf_TroughReleaseBall", 20, Null
'
'*****************************
Function Glf_TroughReleaseBall(args)

    If Not IsNull(args) Then
        If IsObject(args(1)) Then
            If args(1)("new_ball") = True Then
                glf_BIP = glf_BIP + 1
                DispatchPinEvent GLF_BALL_STARTED, Null
                If useBcp Then
                    bcpController.SendPlayerVariable GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL), GetPlayerState(GLF_CURRENT_BALL)-1
                    bcpController.SendPlayerVariable GLF_SCORE, GetPlayerState(GLF_SCORE), GetPlayerState(GLF_SCORE)
                End If
            End If
        End If
    End If
    Glf_WriteDebugLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    glf_plunger.AddIncomingBalls = 1
    swTrough1.kick 90, 10
    DispatchPinEvent "trough_eject", Null
    Glf_WriteDebugLog "Release Ball", "Just Kicked"
    If Not IsNull(args) Then
		If IsObject(args(1)) Then
			Set Glf_TroughReleaseBall = args(1)
		Else
			Glf_TroughReleaseBall = args(1)
		End If
	Else
		Glf_TroughReleaseBall = Null
	End If
End Function

'****************************
' Ball Drain
' Event Listeners:      
    AddPinEventListener GLF_BALL_DRAIN, "ball_drain", "Glf_Drain", 20, Null
'
'*****************************
Function Glf_Drain(args)
    
    If Not glf_gameTilted And glf_gameStarted = True Then
        Dim ballsToSave : ballsToSave = args(1) 
        Glf_WriteDebugLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
        Glf_WriteDebugLog "end_of_ball, balls in play", CStr(glf_BIP)
        If ballsToSave <= 0 Then
            Exit Function
        End If

        glf_BIP = glf_BIP - 1
        glf_debugLog.WriteToLog "Trough", "Ball Drained: BIP: " & glf_BIP

        If glf_BIP > 0 Then
            Exit Function
        End If
        
        DispatchPinEvent GLF_BALL_WILL_END, Null
        DispatchQueuePinEvent GLF_BALL_ENDING, Null
    End If
    
End Function

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener GLF_BALL_ENDING, "ball_will_end", "Glf_BallWillEnd", 10, Null
'
'*****************************
Function Glf_BallWillEnd(args)
    DispatchPinEvent GLF_BALL_ENDED, Null
    If Not IsNull(args) Then
		If IsObject(args(1)) Then
			Set Glf_BallWillEnd = args(1)
		Else
			Glf_BallWillEnd = args(1)
		End If
	Else
		Glf_BallWillEnd = Null
	End If
End Function

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener GLF_BALL_ENDED, "end_of_ball", "Glf_EndOfBall", 20, Null
'
'*****************************
Function Glf_EndOfBall(args)

    If GetPlayerState("extra_balls") > 0 Then
        'self.debug_log("Awarded extra ball to Player %s. Shoot Again", self.player.index + 1)
        'self.player.extra_balls -= 1
        SetPlayerState "extra_balls", GetPlayerState("extra_balls") - 1
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000
        Exit Function
    End If
    
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
    
    If GetPlayerState(GLF_CURRENT_BALL) > glf_game.BallsPerGame() Then
        Dim device
        For Each device in glf_ball_devices.Items()
            If device.HasBall() Then
                device.EjectAll()
            End If
        Next
        DispatchPinEvent "game_will_end", Null
        DispatchQueuePinEvent "game_ending", Null
    Else
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000 
    End If
    


End Function

'****************************
' Game Over
' Event Listeners:      
AddPinEventListener "game_ending", "glf_game_over", "Glf_EndGame", 20, Null
'
'*****************************
Function Glf_EndGame(args)
    If GetPlayerStateForPlayer("0", "score") = False Then
        glf_machine_vars("player1_score").Value = 0
    Else
        glf_machine_vars("player1_score").Value = GetPlayerStateForPlayer("0", "score")
    End If
    If GetPlayerStateForPlayer("1", "score") = False Then
        glf_machine_vars("player2_score").Value = 0
    Else
        glf_machine_vars("player2_score").Value = GetPlayerStateForPlayer("1", "score")
    End If
    If GetPlayerStateForPlayer("2", "score") = False Then
        glf_machine_vars("player3_score").Value = 0
    Else
        glf_machine_vars("player3_score").Value = GetPlayerStateForPlayer("2", "score")
    End If
    If GetPlayerStateForPlayer("3", "score") = False Then
        glf_machine_vars("player4_score").Value = 0
    Else
        glf_machine_vars("player4_score").Value = GetPlayerStateForPlayer("3", "score")
    End If
    glf_gameStarted = False
    glf_currentPlayer = Null
    glf_playerState.RemoveAll()

    DispatchPinEvent "game_ended", Null
    If Not IsNull(args) Then
		If IsObject(args(1)) Then
			Set Glf_EndGame = args(1)
		Else
			Glf_EndGame = args(1)
		End If
	Else
		Glf_EndGame = Null
	End If
End Function

AddPinEventListener "glf_game_cancel", "glf_game_cancel", "Glf_GameCancel", 20, Null

Function Glf_GameCancel(args)
    Dim device
    For Each device in glf_ball_devices.Items()
        If device.HasBall() Then
            device.EjectAll()
        End If
    Next
    Dim mode
    For Each mode in glf_modes.Items()
        mode.StopMode()
    Next
    Dim flipper
    For Each flipper in glf_flippers.Items()
        flipper.Disable()
    Next
    Dim auto_fire_device
    For Each auto_fire_device in glf_autofiredevices.Items()
        auto_fire_device.Disable()
    Next
    glf_bip = 0
    Glf_EndGame Null
    Glf_Reset(Null)
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
        Dim timestamp : timestamp = GetTimeStamp(True)
        Filename = cGameName + "_" & timestamp & "_debug_log.txt"
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
	
	Private Function GetTimeStamp(full)
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
        If full = True Then
            GetTimeStamp = _
            LZ(Year(CurrTime),   4) & "-" _
            & LZ(Month(CurrTime),  2) & "-" _
            & LZ(Day(CurrTime),	2) & "_" _
            & LZ(Hour(CurrTime),   2) & "_" _
            & LZ(Minute(CurrTime), 2) & "_" _
            & LZ(Second(CurrTime), 2) & "_" _
            & LZ(MilliSecs, 4)
        Else
            GetTimeStamp = _
            LZ(Hour(CurrTime),   2) & "_" _
            & LZ(Minute(CurrTime), 2) & "_" _
            & LZ(Second(CurrTime), 2)
        End If
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message)
		If glf_debugEnabled = True Then
			Dim FormattedMsg, Timestamp
			Timestamp = GetTimeStamp(False)
			FormattedMsg = Timestamp & ": " & label & ": " & message
			TxtFileStream.WriteLine FormattedMsg
			Debug.Print label & ": " & message
		End If
		Glf_MonitorEventStream label, message
	End Sub
End Class

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************


Dim glf_lastPinEvent : glf_lastPinEvent = Null

Dim glf_dispatch_parent : glf_dispatch_parent = 0
Dim glf_dispatch_q : Set glf_dispatch_q = CreateObject("Scripting.Dictionary")

Dim glf_frame_dispatch_count : glf_frame_dispatch_count = 0
Dim glf_frame_handler_count : glf_frame_handler_count = 0

Dim glf_dispatch_queue_int : glf_dispatch_queue_int = 0

Sub DispatchPinEvent(e, kwargs)
    AddToDispatchEvents e, kwargs, 1
End Sub

Sub AddToDispatchEvents(e, kwargs, event_type)
    If glf_dispatch_await.Exists(e) Then
        glf_dispatch_await.Remove e
    End If
    glf_dispatch_await.Add e & ";" & glf_dispatch_queue_int, Array(kwargs, event_type)
    glf_dispatch_queue_int = glf_dispatch_queue_int + 1
End Sub

Function DispatchPinHandlers(e, args)
    DispatchPinHandlers = Empty
    Dim handler : handler = args(0)
    Dim event_args, retArgs
    event_args = args(1)
    glf_frame_handler_count = glf_frame_handler_count + 1
    If IsNull(event_args(0)) Then
        Set retArgs = GlfKwargs()
    Else
        If IsObject(event_args(0)) Then
            Set retArgs = event_args(0)
        Else
            retArgs = event_args(0)
        End If
    End If
    If IsObject(retArgs) Then
        Set glf_dispatch_current_kwargs = retArgs	
    Else
        glf_dispatch_current_kwargs = retArgs
    End If
    If event_args(1) = 2 Then 'Queue Event
        Set retArgs = GetRef(handler(0))(Array(handler(2), retArgs, args(2)))
        If IsObject(retArgs) Then
            If retArgs.Exists("wait_for") Then
                DispatchPinHandlers = retArgs("wait_for")
            End If
        End If
    Else
        GetRef(handler(0))(Array(handler(2), event_args(0), args(2)))
    End If
End Function

Sub RunDispatchPinEvent(eKey, kwargs)
    Dim e
    Dim split_key : split_key = Split(eKey, ";")
    e=split_key(0)
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If

    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If

    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            glf_frame_dispatch_count = glf_frame_dispatch_count + 1
            glf_dispatch_handlers_await.Add e&"_"&k(1), Array(handler, kwargs, e)
        Else
            Glf_WriteDebugLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
End Sub

Sub RunAutoFireDispatchPinEvent(e, kwargs)

    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If

    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            glf_frame_dispatch_count = glf_frame_dispatch_count + 1
            'debug.print "Adding Handler for: " & e&"_"&k(1)
            'glf_dispatch_handlers_await.Add e&"_"&k(1), Array(handler, kwargs, e)
            'If SwitchHandler(handler(0), Array(handler(2), kwargs, e)) = False Then
                'debug.print e&"_"&k(1)
                GetRef(handler(0))(Array(handler(2), kwargs, e))
            'End If
        Else
            Glf_WriteDebugLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
    Glf_EventBlocks(e).RemoveAll

End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchReplayPinEvent", e
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        Dim handlers_a1 : handlers_a1 = handlers(k(1))
        kwargs = GetRef(handlers_a1(0))(Array(handlers_a1(2), kwargs, e))
    Next
    DispatchRelayPinEvent = kwargs
    Glf_EventBlocks(e).RemoveAll
End Function

Function DispatchQueuePinEvent(e, kwargs)
    
    AddToDispatchEvents e, kwargs, 2
    Exit Function
    
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchQueuePinEvent", e & " has no listeners"
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
    Glf_WriteDebugLog "DispatchQueuePinEvent", e
    Dim glf_dis_events : glf_dis_events = glf_pinEventsOrder(e)
    For i=0 to UBound(glf_dis_events)
        k = glf_dis_events(i)
        Glf_WriteDebugLog "DispatchQueuePinEvent"&e, "key: " & k(1) & ", priority: " & k(0)
        'Call the handlers.
        'The handlers might return a waitfor command.
        'If NO wait for command, continue calling handlers.
        'IF wait for command, then AddPinEventListener for the waitfor event. The callback handler needs to be ContinueDispatchQueuePinEvent.
        Dim handlers_a1 : handlers_a1 = handlers(k(1))
        Set retArgs = GetRef(handlers_a1(0))(Array(handlers_a1(2), kwargs, e))
        If retArgs.Exists("wait_for") And i<Ubound(glf_dis_events) Then
            'pause execution of handlers at index I. 
            Glf_WriteDebugLog "DispatchQueuePinEvent"&e, k(1) & "_wait_for"
            Dim wait_for : wait_for = retArgs("wait_for")
            kwargs.Remove "wait_for" 
            AddPinEventListener wait_for, k(1) & "_wait_for", "ContinueDispatchQueuePinEvent", k(0), Array(e, kwargs, i+1)
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
    
    Dim ownProps : ownProps = args(0)
    Dim kwargs
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1)
    End If


    Dim i,key,keys,items
    keys=ownProps(0)
    items=ownProps(1)

    'Inject handlers back into dispatch
    For i=0 to UBound(keys)
        glf_dispatch_handlers_await.Add keys(i), items(i)
    Next
    Exit Function


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
                Dim temp_a1 : temp_a1 = tempArray(i)
                If priority > temp_a1(0) Then ' Compare priorities
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
                Dim temp_a2 : temp_a2 = tempArray(i)
                If temp_a2(1) = key Then
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
        GetPlayerState = False
    End If
End Function

Function GetPlayerStateForPlayer(player, key)
    dim p
    Select Case player
        Case 0:
            p = "PLAYER 1"
        Case 1:
            p = "PLAYER 2"
        Case 2:
            p = "PLAYER 3"
        Case 3:
            p = "PLAYER 4"
    End Select

    If glf_playerState.Exists(p) Then
        GetPlayerStateForPlayer = glf_playerState(p)(key)
    Else
        GetPlayerStateForPlayer = False
    End If
End Function

Function Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            Getglf_currentPlayerNumber = 0
        Case "PLAYER 2":
            Getglf_currentPlayerNumber = 1
        Case "PLAYER 3":
            Getglf_currentPlayerNumber = 2
        Case "PLAYER 4":
            Getglf_currentPlayerNumber = 3
    End Select
End Function

Function Getglf_PlayerName(player)
    Select Case player
        Case 0:
            Getglf_PlayerName = "PLAYER 1"
        Case 1:
            Getglf_PlayerName = "PLAYER 2"
        Case 2:
            Getglf_PlayerName = "PLAYER 3"
        Case 3:
            Getglf_PlayerName = "PLAYER 4"
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
        If VarType(GetPlayerState(key)) <> vbBoolean Then
            If GetPlayerState(key) = value Then
                Exit Function
            End If
        End If
    End If   
    Dim prevValue
    If glf_playerState(glf_currentPlayer).Exists(key) Then
        prevValue = glf_playerState(glf_currentPlayer)(key)
        glf_playerState(glf_currentPlayer).Remove key
    End If
    glf_playerState(glf_currentPlayer).Add key, value
    If glf_debug_level = "Debug" Then
        Dim p,v : p = prevValue : v = value
        If IsNull(prevValue) Then
            p=""
        End If
        If IsNull(value) Then
            v=""
        End If
        Glf_WriteDebugLog "Player State", "Variable "& key &" changed from " & CStr(p) & " to " & CStr(v)
    End If
    Glf_MonitorPlayerStateUpdate key, value
    If glf_monitor_player_vars = True Then
        Glf_BcpSendPlayerVar Array(Null, Array(key, value, prevValue))
    End If
    If glf_playerEvents.Exists(key) Then
        SetDelay "FirePlayerEventHandlers_" & key, "FirePlayerEventHandlersProxy",  Array(key, value, prevValue, -1), 200
        'FirePlayerEventHandlers key, value, prevValue, -1
    End If
    
    SetPlayerState = Null
End Function

Function SetPlayerStateByPlayer(key, value, player)

    If IsArray(value) Then
        If Join(GetPlayerStateForPlayer(player, key)) = Join(value) Then
            Exit Function
        End If
    Else
        If GetPlayerStateForPlayer(player, key) = value Then
            Exit Function
        End If
    End If   
    Dim prevValue, player_name
    player_name = Getglf_PlayerName(player)
    If glf_playerState(player_name).Exists(key) Then
        prevValue = glf_playerState(player_name)(key)
        glf_playerState(player_name).Remove key
    End If
    glf_playerState(player_name).Add key, value
    
    If glf_playerEvents.Exists(key) Then
        SetDelay "FirePlayerEventHandlers_" & key, "FirePlayerEventHandlersProxy",  Array(key, value, prevValue, player), 200
        'FirePlayerEventHandlers key, value, prevValue, player
    End If
    
    SetPlayerStateByPlayer = Null
End Function

Sub FirePlayerEventHandlersProxy(args)
    FirePlayerEventHandlers args(0), args(1), args(2), args(3)
End Sub

Sub FirePlayerEventHandlers(evt, value, prevValue, player)
    If Not glf_playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = glf_playerEvents(evt)
    For Each k In glf_playerEventsOrder(evt)
        Dim handlers_a1 : handlers_a1 = handlers(k(1))
        If handlers_a1(3) = player or handlers_a1(3) = Getglf_currentPlayerNumber() Then
            GetRef(handlers_a1(0))(Array(handlers_a1(2), Array(evt,value,prevValue)))
        End If
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, player, callbackName, priority, args)
    If Not glf_playerEvents.Exists(evt) Then
        glf_playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_playerEvents(evt).Exists(key) Then
        glf_playerEvents(evt).Add key, Array(callbackName, priority, args, player)
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
                Dim temp_a3 : temp_a3 = tempArray(i)
                If priority > temp_a3(0) Then ' Compare priorities
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
                Dim temp_a4 : temp_a4 = tempArray(i)
                If temp_a4(1) = key Then
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
        FirePlayerEventHandlers key, GetPlayerState(key), GetPlayerState(key), -1
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
