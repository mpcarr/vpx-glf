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

