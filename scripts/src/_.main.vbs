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
Dim glf_ball_holds : Set glf_ball_holds = CreateObject("Scripting.Dictionary")
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

Public Function Glf_ParseInput(value, isTime)
	Dim templateCode : templateCode = ""
	Dim tmp: tmp = value
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
	value = Replace(value, "&&", "And")
	Glf_ConvertCondition = "    "&retName&" = " & value
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

	Public default Function init(input, isTime)
        m_raw = input
        Dim parsedInput : parsedInput = Glf_ParseInput(input, isTime)
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

