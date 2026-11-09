' Run with: cscript //nologo tests\ball-device-tests.vbs
' Runs the source class with GLF's real delay/event helpers and frame dispatcher.
Option Explicit
Dim fso, root, testDir, gameTime
Set fso = CreateObject("Scripting.FileSystemObject")
testDir = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(testDir)
Dim glf_ball_devices, glf_plunger, glf_debugLog
Dim glf_pinEvents, glf_pinEventsOrder, Glf_EventBlocks
Dim glf_dispatch_await, glf_dispatch_handlers_await, glf_dispatch_current_kwargs
Dim glf_last_switch_hit, glf_last_switch_hit_time, glf_max_dispatch
Dim gKicks, gDisables, gLastBall, gEvents, gClaims, gClaimResult, gBlock, gBlockSecond
Dim gAssertions, gCases, gCase, gToken, gFailures
gAssertions = 0: gCases = 0: gFailures = 0
LoadScript root & "\src\delay.vbs"
LoadScript root & "\src\pin-event-handlers.vbs"
LoadFrameDispatcher root & "\src\_.main.vbs"
LoadScript root & "\src\devices\ball-device.vbs"

Sub LoadScript(path)
    Dim stream: Set stream = fso.OpenTextFile(path)
    ExecuteGlobal stream.ReadAll
    stream.Close
End Sub
' Load only the real frame dispatcher; _.main.vbs also contains VPX-only setup.
Sub LoadFrameDispatcher(path)
    Dim stream, source, startAt, endAt
    Set stream = fso.OpenTextFile(path)
    source = stream.ReadAll
    stream.Close
    startAt = InStr(1, source, "Public Function Glf_RunHandlers(i)", vbTextCompare)
    If startAt = 0 Then Err.Raise 5, "ball-device-tests", "Cannot find Glf_RunHandlers in " & path
    endAt = InStr(startAt, source, "End Function", vbTextCompare)
    If endAt = 0 Then Err.Raise 5, "ball-device-tests", "Cannot find the end of Glf_RunHandlers"
    ExecuteGlobal Mid(source, startAt, endAt - startAt + Len("End Function"))
End Sub
Class TestBall
    Public ID
End Class
Function Ball(id)
    Dim b: Set b = New TestBall
    b.ID = id
    Set Ball = b
End Function
Sub Glf_WriteDebugLog(a, b)
End Sub
Function GlfKwargs()
    Set GlfKwargs = CreateObject("Scripting.Dictionary")
End Function
Sub Kick(ball)
    If IsObject(ball) Then
        gKicks = gKicks + 1
        gLastBall = ball.ID
    Else
        gDisables = gDisables + 1
    End If
End Sub
Function Observe(args)
    Dim name: name = args(0)
    gEvents(name) = Events(name) + 1
    If IsObject(args(1)) Then
        ' Ordinary/raw switch observations carry a ball, not event kwargs.
        If TypeName(args(1)) = "Dictionary" Then
            If args(1).Exists("glf_eject_token") Then gToken = args(1)("glf_eject_token")
        End If
        Set Observe = args(1)
    Else
        Observe = args(1)
    End If
End Function
Function Claim(args)
    gClaims = gClaims + 1
    Claim = gClaimResult
End Function
Function BlockAttempt(args)
    If gBlock Then args(1).Add "wait_for", "allow_kick"
    Set BlockAttempt = args(1)
End Function
Function BlockSecondAttempt(args)
    If gBlockSecond Then args(1).Add "wait_for", "allow_second_kick"
    Set BlockSecondAttempt = args(1)
End Function
Function Events(name)
    Events = 0
    If gEvents.Exists(name) Then Events = gEvents(name)
End Function
Sub Check(actual, expected, description)
    gAssertions = gAssertions + 1
    If actual <> expected Then
        WScript.Echo "FAIL: " & gCase & " / " & description & " (actual=" & actual & ", expected=" & expected & ")"
        gFailures = gFailures + 1
    End If
End Sub
Function Device(caseName, switchCount)
    gCase = caseName: gCases = gCases + 1
    gameTime = 1000
    gKicks = 0: gDisables = 0: gLastBall = -1: gClaims = 0: gClaimResult = 1: gBlock = False
    gBlockSecond = False
    Set gEvents = CreateObject("Scripting.Dictionary")
    Set glf_ball_devices = CreateObject("Scripting.Dictionary")
    Set glf_pinEvents = CreateObject("Scripting.Dictionary")
    Set glf_pinEventsOrder = CreateObject("Scripting.Dictionary")
    Set Glf_EventBlocks = CreateObject("Scripting.Dictionary")
    Set glf_dispatch_await = CreateObject("Scripting.Dictionary")
    Set glf_dispatch_handlers_await = CreateObject("Scripting.Dictionary")
    Set delayQueue = CreateObject("Scripting.Dictionary")
    Set delayQueueMap = CreateObject("Scripting.Dictionary")
    Set delayCallbacks = CreateObject("Scripting.Dictionary")
    glf_dispatch_queue_int = 0
    glf_last_switch_hit = "": glf_last_switch_hit_time = 0
    glf_max_dispatch = 1000
    Dim d, switches(), i, evt
    Set d = CreateGlfBallDevice("test")
    If switchCount > 0 Then
        ReDim switches(switchCount - 1)
        For i = 0 To switchCount - 1: switches(i) = "s" & i: Next
        d.BallSwitches = switches
    End If
    d.EntranceCountDelay = 20
    d.ExitCountDelay = 20
    d.EjectTimeout = 100
    d.EjectRetryDelay = 30
    d.BallMissingTimeout = 200
    d.IdleMissingBallTimeout = 100
    d.AutoFireOnUnexpectedBall = False
    d.EjectCallback = "Kick"
    For Each evt In Array("ball_entered", "ball_exiting", "ball_eject_attempt", "ejecting_ball", _
            "ball_eject_success", "ball_eject_failed", "broken", "ball_missing", "ball_count_changed")
        AddPinEventListener d.Name & "_" & evt, "observe", "Observe", 100, evt
    Next
    AddPinEventListener d.Name & "_ball_enter", "claim", "Claim", 100, Null
    Set Device = d
End Function
Sub Frame()
    Dim i, key, keys
    i = Glf_RunHandlers(0)
    keys = glf_dispatch_await.Keys
    For Each key In keys
        RunDispatchPinEvent key, glf_dispatch_await(key)
        glf_dispatch_await.Remove key
        i = Glf_RunHandlers(i)
    Next
    DelayTick
End Sub
Sub Advance(ms)
    Dim i
    For i = 1 To ms
        gameTime = gameTime + 1
        Frame
    Next
End Sub
Sub Hit(b, slot)
    TableHit b, "s" & slot
End Sub
Sub TableHit(b, switchName)
    ' Match the generated VPX _Hit handler, including the shared debounce clock.
    DispatchPinEvent switchName & "_active", b
    glf_last_switch_hit_time = gameTime
    glf_last_switch_hit = switchName
    Advance 2
End Sub
Sub Unhit(b, slot)
    DispatchPinEvent "s" & slot & "_inactive", b
    Advance 2
End Sub
Sub Confirm(b)
    TableHit b, "exit"
End Sub

Dim d, a, b, c, oldToken, i, j, slot, which, physical(3, 2), balls(3), expected, occupied
Set a = Ball(1): Set b = Ball(2): Set c = Ball(3)

Set d = Device("empty unconfigured device", 0)
Check d.HasBall, False, "HasBall is safe"
Check d.Balls, 0, "zero balls"
d.EjectAll: d.EjectBalls 0: Advance 10
Check d.PendingEjects, 0, "no request for zero balls"

Set d = Device("debounce arrival and departure", 1)
Hit a, 0
Check d.Balls, 0, "arrival waits"
Advance 25: Check d.Balls, 1, "settled arrival"
Unhit a, 0: Check d.Balls, 1, "departure waits"
Advance 25: Check d.Balls, 0, "settled departure"

Set d = Device("short entrance ignored", 1)
Hit a, 0: Unhit a, 0: Advance 30
Check d.Balls, 0, "no ball"
Check Events("ball_entered"), 0, "no capture"

Set d = Device("duplicate edge events", 1)
Hit a, 0: Advance 30: Hit a, 0: Advance 30
Check d.Balls, 1, "duplicate hit does not increment"
Check Events("ball_entered"), 1, "one capture"
Unhit a, 0: Advance 30: Unhit a, 0: Advance 30
Check d.Balls, 0, "duplicate unhit cannot make negative"
Check Events("ball_exiting"), 1, "one departure"

Set d = Device("same ball overlaps switches", 3)
Hit a, 0: Hit a, 1: Advance 30
Check d.Balls, 1, "unique ball count"
Unhit a, 0: Advance 30
Check d.Balls, 1, "still in device"
Check Events("ball_exiting"), 0, "not an exit"
Check Events("ball_entered"), 1, "one capture"

Set d = Device("switch gap during movement", 3)
Hit a, 0: Advance 30: Unhit a, 0: Advance 5: Hit a, 2: Advance 30
Check d.Balls, 1, "movement preserves count"
Check Events("ball_entered"), 1, "not a second capture"
Check Events("ball_exiting"), 0, "not a departure"
Check d.HasBall, True, "ball away from eject position is visible"

Set d = Device("multiple different balls", 3)
Hit a, 0: Hit b, 1: Hit c, 2: Advance 30
Check d.Balls, 3, "three balls"
Unhit b, 1: Advance 30
Check d.Balls, 2, "middle ball departure"

Set d = Device("two balls overlap one trigger", 1)
Hit a, 0: Hit b, 0: Advance 30: Unhit a, 0: Advance 30
Check d.Balls, 1, "remaining ball stays counted"

Set d = Device("failed kick retries", 1)
Hit a, 0: Advance 30: d.Eject: Advance 10
Check gKicks, 1, "initial kick"
Advance 140
Check gKicks, 2, "retry"
Check Events("ball_eject_success"), 0, "stationary ball never succeeds"
Check Events("ball_eject_failed"), 1, "failure event"

Set d = Device("retry limit", 1)
d.MaxEjectAttempts = 3
Hit a, 0: Advance 30: d.Eject: Advance 500
Check gKicks, 3, "exact attempt limit"
Check d.State, "eject_broken", "broken state"
Check Events("broken"), 1, "one broken event"
Advance 300: Check gKicks, 3, "stays stopped"

Set d = Device("rapid lock bounce retries with final ball still present", 3)
d.EjectTargets = Array("exit")
Hit a, 2: Unhit a, 2: Hit a, 1: Unhit a, 1: Hit a, 0: Advance 60
d.Eject: Advance 10
Unhit a, 0: Hit a, 0: Unhit a, 0: Hit a, 0: Unhit a, 0: Hit a, 0
Advance 140
Check d.Balls, 1, "final re-hit survives the 50ms dispatch filter"
Check gKicks, 2, "same ball is retried after failed departure"
Check d.PendingEjects, 1, "same request stays reserved"
Check gClaims, 1, "bounce is not another capture"
Check Events("ball_eject_success"), 0, "no success for a ball still in the lock"
Check Events("ball_eject_failed"), 1, "one failed first attempt"

Set d = Device("ordinary listeners retain hit debounce", 1)
AddPinEventListener "s0_active", "normal_hit", "Observe", 100, "normal_hit"
Hit a, 0: Advance 30: Unhit a, 0: Hit a, 0: Advance 30
Check Events("normal_hit"), 1, "scoring listeners do not get the rapid re-hit"
Check d.Balls, 1, "device still gets the final active edge"
Check Events("ball_entered"), 1, "bounce does not duplicate a capture"
Advance 60: Hit a, 0
Check Events("normal_hit"), 2, "later ordinary hit is still delivered"

Set d = Device("two balls on one sensor inside debounce", 1)
Hit a, 0: Hit b, 0: Advance 30
Check d.Balls, 2, "both identities counted despite shared switch debounce"
Unhit a, 0: Advance 30
Check d.Balls, 1, "second identity remains after first leaves"

Set d = Device("queued raw edges preserve order", 1)
DispatchPinEvent "s0_active", a
glf_last_switch_hit_time = gameTime: glf_last_switch_hit = "s0"
DispatchPinEvent "s0_inactive", a
DispatchPinEvent "s0_active", a
Advance 30
Check d.Balls, 1, "active inactive active in one frame ends occupied"
Check Events("ball_entered"), 1, "one settled capture"
Check Events("ball_exiting"), 0, "no transient settled departure"

Set d = Device("old occupancy listeners removed on reconfiguration", 1)
d.BallSwitches = Array("new_sensor")
TableHit a, "s0": Advance 30
Check d.Balls, 0, "old switch no longer changes occupancy"
TableHit a, "new_sensor": Advance 30
Check d.Balls, 1, "replacement switch tracks occupancy"

Set d = Device("count-confirmed successful eject", 1)
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 140
Check Events("ball_eject_success"), 1, "one success"
Check d.PendingEjects, 0, "request consumed"
Advance 300: Check gKicks, 1, "no residual retry"

Set d = Device("returned ball retries original request", 1)
d.AddIncomingBalls = 1
Hit a, 0: Advance 30
Check d.IncomingBalls, 0, "incoming expectation consumed on arrival"
d.Eject: Advance 10: Unhit a, 0: Advance 30: Hit a, 0: Advance 70
Check gKicks, 2, "returned expected ball retried"
Check Events("ball_entered"), 1, "return not a new arrival"
Check gClaims, 0, "return not offered to lock"
Check Events("ball_eject_success"), 0, "no premature success"

Set d = Device("return with automatic capture enabled", 1)
d.AutoFireOnUnexpectedBall = True
Hit a, 0: Advance 540: Unhit a, 0: Advance 30: Hit a, 0: Advance 70
Check d.PendingEjects, 1, "return does not enqueue another auto eject"
Check gClaims, 1, "only first arrival claimed"
Check gKicks, 2, "one retry"

Set d = Device("brief return crossing timeout", 1)
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 100: Hit a, 0: Advance 10
Check Events("ball_eject_success"), 0, "raw return blocks pending timeout"
Advance 50: Check gKicks, 2, "retry after settle"

Set d = Device("incoming ball during another eject", 2)
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Hit b, 1: Advance 140
Check d.Balls, 1, "new ball retained"
Check Events("ball_entered"), 2, "new identity recognised"
Check Events("ball_eject_success"), 1, "original eject completes"

Set d = Device("wrong confirmation ignored", 1)
d.EjectTargets = Array("exit")
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 30: Confirm b
Check Events("ball_eject_success"), 0, "different ball cannot confirm"
Confirm a: Advance 10
Check Events("ball_eject_success"), 1, "matching ball confirms"

Set d = Device("matching confirmation inside another ball's debounce", 1)
d.EjectTargets = Array("exit")
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 30
Confirm b: Confirm a: Advance 10
Check Events("ball_eject_success"), 1, "correct ball confirms despite recent wrong-ball hit"
Check d.PendingEjects, 0, "confirmed request retired"
Confirm a: Advance 10
Check Events("ball_eject_success"), 1, "removed raw confirmation listener cannot complete twice"

Set d = Device("confirmation before exit settles", 1)
d.EjectTargets = Array("exit")
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Confirm a
Check Events("ball_eject_success"), 0, "confirmation waits for count"
Advance 30: Check Events("ball_eject_success"), 1, "latched confirmation completes"
Advance 400: Check Events("ball_eject_success"), 1, "old timeout cannot duplicate"

Set d = Device("confirmation while still present", 1)
d.EjectTargets = Array("exit")
Hit a, 0: Advance 30: d.Eject: Advance 10: Confirm a
Check Events("ball_eject_success"), 0, "departure required"

Set d = Device("idle confirmation ignored", 1)
d.EjectTargets = Array("exit")
Confirm a: d.BallExitSuccess a: Advance 10
Check Events("ball_eject_success"), 0, "no active attempt"

Set d = Device("late explicit confirmation", 1)
d.EjectTargets = Array("exit")
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 160
Check Events("ball_eject_success"), 0, "explicit target required"
Confirm a: Advance 250
Check Events("ball_eject_success"), 1, "late confirm accepted"
Check Events("ball_missing"), 0, "missing timer cancelled"

Set d = Device("missing confirmation does not eject another ball", 2)
d.EjectTargets = Array("exit")
Hit a, 0: Hit b, 1: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 360
Check Events("ball_missing"), 1, "reports missing"
Check Events("ball_eject_success"), 0, "does not fake success"
Check d.PendingEjects, 0, "retires missing request"
Check gKicks, 1, "does not eject replacement"

Set d = Device("overlapping explicit eject requests", 2)
Hit a, 0: Hit b, 1: Advance 30: d.Eject: d.Eject: Advance 10
Check gKicks, 1, "only first request fires"
Check d.PendingEjects, 2, "both requests retained"
Unhit a, 0: Advance 140
Check gKicks, 1, "next ball not yet at eject position"
Unhit b, 1: Hit b, 0: Advance 40
Check gKicks, 2, "second request fires when ready"
Unhit b, 0: Advance 140
Check d.PendingEjects, 0, "both completed"
Advance 300: Check gKicks, 2, "no extra request"

Set d = Device("eject-all idempotent and late seating", 2)
Hit a, 0: Hit b, 1: Advance 30: d.EjectAll: d.EjectAll: Advance 10
Check d.PendingEjects, 2, "does not reserve twice"
Unhit a, 0: Advance 140: Unhit b, 1: Hit b, 0: Advance 40: Unhit b, 0: Advance 140
Check gKicks, 2, "exact number ejected"
Check d.PendingEjects, 0, "no counter leak"
d.EjectAll: Advance 50: Check d.PendingEjects, 0, "empty eject all is no-op"

Set d = Device("eject-all while first ball awaits confirmation", 3)
Hit a, 0: Hit b, 1: Hit c, 2: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 30
d.EjectAll: d.EjectAll
Check d.PendingEjects, 3, "includes two remaining balls plus departing ball"

Set d = Device("request while empty", 1)
d.Eject: Advance 60
Check gKicks, 0, "no kick while empty"
Hit a, 0: Advance 40: Check gKicks, 1, "queued request resumes on arrival"

Set d = Device("explicit request consumes automatic arrival", 1)
d.AutoFireOnUnexpectedBall = True
d.Eject: Advance 20: Hit a, 0: Advance 40
Check d.PendingEjects, 1, "one reservation for the arriving ball"
Unhit a, 0: Advance 650
Check gKicks, 1, "no second automatic kick"
Check d.PendingEjects, 0, "no orphan automatic request"

Set d = Device("one explicit request and two automatic arrivals", 2)
d.AutoFireOnUnexpectedBall = True
d.Eject: Advance 20: Hit a, 0: Hit b, 1: Advance 40
Check d.PendingEjects, 2, "two distinct balls reserved once each"
Unhit a, 0: Advance 140: Unhit b, 1: Hit b, 0: Advance 350
Check gKicks, 2, "second ball retains its request"

Set d = Device("unexpected rear ball releases the front locked ball", 3)
d.AutoFireOnUnexpectedBall = True: d.EjectTargets = Array("exit")
gClaimResult = 0: Hit a, 0: Advance 30
Check d.Balls, 1, "one claimed ball held at eject position"
Check d.PendingEjects, 0, "claimed arrival does not request a release"
gClaimResult = 1: Hit b, 1: Advance 540
Check d.Balls, 2, "both balls counted before departure"
Check gKicks, 1, "unexpected arrival requests one kick"
Check gLastBall, a.ID, "front ball chosen instead of blocked rear arrival"
Check d.PendingEjects, 1, "only one release reserved"
Unhit a, 0: Advance 30: Confirm a: Advance 10
Check Events("ball_eject_success"), 1, "front ball confirms the release"
Check d.Balls, 1, "one physical ball remains locked"
Check d.PendingEjects, 0, "release request completed"
Unhit b, 1: Hit b, 0: Advance 600
Check gKicks, 1, "remaining ball is retained after moving forward"
Check gClaims, 2, "each arrival offered to claim relay once"

Set d = Device("unexpected rear arrival retries the selected front ball", 2)
d.AutoFireOnUnexpectedBall = True: d.EjectTargets = Array("exit")
gClaimResult = 0: Hit a, 0: Advance 30
gClaimResult = 1: Hit b, 1: Advance 540
Unhit a, 0: Advance 30: Hit a, 0: Advance 70
Check gKicks, 2, "returned front ball is retried"
Check gLastBall, a.ID, "retry remains bound to actual ejected ball"
Check d.PendingEjects, 1, "return does not add a second release"
Check gClaims, 2, "returned front ball is not a new lock capture"
Unhit a, 0: Advance 30: Confirm b
Check Events("ball_eject_success"), 0, "rear arrival cannot confirm front eject"
Confirm a: Advance 10
Check d.Balls, 1, "one retained after retry succeeds"
Check d.PendingEjects, 0, "retried request consumed once"

Set d = Device("two unexpected arrivals retain one claimed ball", 3)
d.AutoFireOnUnexpectedBall = True: d.EjectTargets = Array("exit")
gClaimResult = 0: Hit a, 0: Advance 30
gClaimResult = 1: Hit b, 1: Hit c, 2: Advance 540
Check d.PendingEjects, 2, "one release per unclaimed arrival"
Check gKicks, 1, "only front ball kicked initially"
Check gLastBall, a.ID, "first release selects front ball"
Unhit a, 0: Advance 30: Confirm a: Advance 10
Unhit b, 1: Hit b, 0: Advance 40
Check gKicks, 2, "next reserved ball ejects after moving forward"
Check gLastBall, b.ID, "second release selects next ball"
Unhit b, 0: Advance 30: Confirm b: Advance 10
Unhit c, 2: Hit c, 0: Advance 600
Check d.Balls, 1, "one ball retained after two releases"
Check d.PendingEjects, 0, "both excess-ball releases completed"
Check gKicks, 2, "retained ball is not ejected again"
Check gClaims, 3, "claim relay sees three original arrivals"

Set d = Device("empty launch button does not arm future ball", 1)
d.PlayerControlledEjectEvent = "launch"
DispatchPinEvent "launch", Null: Advance 10: Hit a, 0: Advance 100
Check gKicks, 0, "later ball not auto launched"
DispatchPinEvent "launch", Null: Advance 10: Check gKicks, 1, "occupied button launches"

Set d = Device("queued kick respects wait_for", 1)
gBlock = True
AddPinEventListener d.Name & "_ball_eject_attempt", "block", "BlockAttempt", 1000, Null
Hit a, 0: Advance 30: d.Eject: Advance 200
Check gKicks, 0, "queue blocks callback and timeout"
gBlock = False: DispatchPinEvent "allow_kick", Null: Advance 10
Check gKicks, 1, "clear queue permits callback"

Set d = Device("sequential queue blockers consume their own wait signals", 1)
gBlock = True: gBlockSecond = True
AddPinEventListener d.Name & "_ball_eject_attempt", "block", "BlockAttempt", 1000, Null
AddPinEventListener d.Name & "_ball_eject_attempt", "block_second", "BlockSecondAttempt", 500, Null
Hit a, 0: Advance 30: d.Eject: Advance 20
Check gKicks, 0, "first blocker holds the kick"
gBlock = False: DispatchPinEvent "allow_kick", Null: Advance 20
Check gKicks, 0, "second blocker holds after the first release"
gBlockSecond = False: DispatchPinEvent "allow_second_kick", Null: Advance 10
Check gKicks, 1, "both releases permit exactly one kick"
DispatchPinEvent "allow_kick", Null: DispatchPinEvent "allow_second_kick", Null: Advance 10
Check gKicks, 1, "old release events cannot replay the kick"

Set d = Device("ball moves while queue is blocked", 2)
gBlock = True
AddPinEventListener d.Name & "_ball_eject_attempt", "block", "BlockAttempt", 1000, Null
Hit a, 0: Advance 30: d.Eject: Advance 10: Unhit a, 0: Hit a, 1: Advance 30
gBlock = False: DispatchPinEvent "allow_kick", Null: Advance 10
Check gKicks, 0, "does not kick absent position"
Unhit a, 1: Hit a, 0: Advance 40
Check gKicks, 1, "resumes when ball moves back"
Check Events("ball_entered"), 1, "movement is not a capture"

Set d = Device("ball disappears while queue is blocked", 2)
gBlock = True
AddPinEventListener d.Name & "_ball_eject_attempt", "block", "BlockAttempt", 1000, Null
Hit a, 0: Hit b, 1: Advance 30: d.Eject: Advance 10: Unhit a, 0: Advance 30
gBlock = False: DispatchPinEvent "allow_kick", Null: Advance 10
Unhit b, 1: Hit b, 0: Advance 40
Check gKicks, 0, "does not substitute another ball"
Check d.PendingEjects, 0, "invalid request retired"

Set d = Device("stale attempt callbacks ignored", 1)
Hit a, 0: Advance 30: d.Eject: Advance 10: oldToken = gToken
Advance 140: Check gKicks, 2, "second attempt started"
d.OnTimer "departure_timeout", oldToken
d.OnTimer "confirm_timeout", oldToken
d.ConfirmEject a, oldToken
Advance 5
Check gKicks, 2, "stale timeout does not retry"
Check Events("ball_eject_success"), 0, "stale confirm ignored"

Set d = Device("incoming expectation only consumed once", 2)
d.AddIncomingBalls = 1
Hit a, 0: Advance 30: Hit b, 1: Advance 30
Check d.IncomingBalls, 0, "expectation consumed"
Check gClaims, 1, "second ball is unexpected"

Set d = Device("idle lost identity recovery", 2)
Hit a, 0: Advance 30: Unhit a, 0: Advance 30: Hit b, 1: Advance 30
Check gClaims, 2, "different ball not suppressed"
Hit a, 0: Advance 30
Check gClaims, 2, "same ball recovery not new capture"
Advance 150: Check Events("ball_missing"), 0, "recovery cancels missing report"

Set d = Device("idle missing ball event", 1)
Hit a, 0: Advance 30: Unhit a, 0: Advance 150
Check Events("ball_missing"), 1, "missing reported"
Advance 200: Check Events("ball_missing"), 1, "reported once"

Set d = Device("idle recovery during missing deadline debounce", 1)
Hit a, 0: Advance 30: Unhit a, 0: Advance 105: Hit a, 0: Advance 40
Check Events("ball_missing"), 0, "raw return defers missing report"
Check Events("ball_entered"), 1, "no duplicate capture"

Set d = Device("manual mechanical launch", 1)
d.MechanicalEject = True: d.EjectCallback = Null
Hit a, 0: Advance 30: Unhit a, 0: Advance 140
Check Events("ball_eject_success"), 1, "manual launch tracked"
Check gKicks, 0, "no powered kick"

Set d = Device("manual return waits for player", 1)
d.MechanicalEject = True
Hit a, 0: Advance 30: Unhit a, 0: Advance 30: Hit a, 0: Advance 150
Check d.State, "waiting_for_player", "wait after return"
Check gKicks, 0, "no unwanted auto launch"
d.Eject: Advance 10: Check gKicks, 1, "explicit launch upgrades request"
Check d.PendingEjects, 1, "no duplicate request"

Set d = Device("manual relaunch after return", 1)
d.MechanicalEject = True: d.EjectCallback = Null
Hit a, 0: Advance 30: Unhit a, 0: Advance 30: Hit a, 0: Advance 30
Unhit a, 0: Advance 140
Check Events("ball_eject_success"), 1, "second manual launch succeeds"
Check d.PendingEjects, 0, "same request consumed"

Set d = Device("enabled eject callback duration", 1)
d.EjectEnableTime = 40
Hit a, 0: Advance 30: d.Eject: Advance 10
Check gDisables, 0, "not disabled too soon"
Advance 40: Check gDisables, 1, "disabled once"

Set d = Device("stop invalidates pending callbacks", 1)
d.EjectEnableTime = 100
Hit a, 0: Advance 30: d.Eject: Advance 10: oldToken = gToken: d.StopDevice
Check gDisables, 1, "stop disables active callback"
d.OnTimer "departure_timeout", oldToken: Advance 300
Check gKicks, 1, "no delayed retries"
Check d.PendingEjects, 0, "queue cleared"
Check d.State, "stopped", "stopped state"

Set d = Device("reset broken device", 1)
d.MaxEjectAttempts = 1
Hit a, 0: Advance 30: d.Eject: Advance 140
d.Reset: d.Eject: Advance 10
Check gKicks, 2, "reset allows a new request"

Set d = Device("switches remain tracked while stopped", 1)
Hit a, 0: Advance 30: d.StopDevice: Unhit a, 0: Advance 30: d.Reset: Advance 30
Check d.Balls, 0, "reset recounts changed switch state"

Set d = Device("plunger readiness through complete launch", 1)
d.DefaultDevice = True: d.MechanicalEject = True
Check d.IsReadyToReceive, True, "empty plunger ready"
d.AddIncomingBalls = 1: Check d.IsReadyToReceive, False, "incoming ball reserves space"
Hit a, 0: Advance 30
Check d.IsReadyToReceive, False, "occupied plunger not ready"
Check d.IncomingBalls, 0, "expectation consumed"
d.Eject: Advance 10: Unhit a, 0: Advance 30
Check d.IsReadyToReceive, False, "awaiting launch confirmation"
Advance 110: Check d.IsReadyToReceive, True, "ready after confirmed launch"

Set d = Device("seeded starting balls", 2)
d.SeedBall a, 0: d.SeedBall a, 1
Check d.Balls, 1, "seed deduplicates"
Hit a, 0: Advance 30
Check Events("ball_entered"), 0, "seed not recaptured"

Set d = Device("ball search does not overlap eject", 1)
Hit a, 0: Advance 30: d.Eject: Advance 10: d.BallSearch 3
Check gKicks, 1, "search leaves eject alone"

Set d = Device("empty ball search disable timing", 1)
d.EjectEnableTime = 40: d.BallSearch 1
Check gDisables, 1, "empty search invokes callback with Null"
Advance 50: Check gDisables, 2, "enabled search also receives disable"

Set d = Device("configurable eject switch", 2)
d.EjectSwitchIndex = 1
Hit a, 1: Advance 30: d.Eject: Advance 10
Check gKicks, 1, "uses configured position"
Check gLastBall, 1, "correct physical ball"

Set d = Device("randomized switch identity conservation", 3)
For i = 0 To 3
    Set balls(i) = Ball(i + 10)
    For j = 0 To 2: physical(i, j) = False: Next
Next
Dim randomSeed: randomSeed = Rnd(-7): Randomize 7
For i = 1 To 120
    which = Int(Rnd * 4): slot = Int(Rnd * 3)
    If Rnd < 0.5 Then
        Hit balls(which), slot
        physical(which, slot) = True
    Else
        Unhit balls(which), slot
        physical(which, slot) = False
    End If
    Advance 25
    expected = 0
    For which = 0 To 3
        occupied = False
        For j = 0 To 2
            If physical(which, j) Then occupied = True
        Next
        If occupied Then expected = expected + 1
    Next
    Check d.Balls, expected, "step " & i
Next

WScript.Echo gCases & " scenarios; " & gAssertions & " assertions; " & gFailures & " failures."
If gFailures > 0 Then WScript.Quit 1
WScript.Quit 0
