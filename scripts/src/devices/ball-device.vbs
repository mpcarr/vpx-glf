' GLF ball device: serialized ejects and settled, identity-based ball counting.
' BallSwitches are occupancy switches. VPX ball.ID identifies a physical ball.
' EjectTargets retains GLF's meaning: names whose _active events confirm an eject.
Function CreateGlfBallDevice(name)
    Dim device : Set device = (New GlfBallDevice)(name)
    Set CreateGlfBallDevice = device
End Function

Class GlfBallDevice
    Private m_name, m_ball_switches, m_switch_balls, m_switch_due
    Private m_present, m_lost, m_queue, m_next_request
    Private m_eject_events, m_eject_all_events, m_eject_targets
    Private m_player_controlled_eject_event, m_eject_callback
    Private m_eject_timeout, m_eject_enable_time, m_entrance_count_delay, m_exit_count_delay
    Private m_retry_delay, m_max_attempts, m_missing_timeout, m_idle_missing_timeout
    Private m_eject_switch, m_incoming_balls, m_default_device, m_mechanical_eject
    Private m_auto_fire_on_unexpected_ball, m_exclude_from_ball_search, m_debug
    Private m_state, m_request, m_attempt, m_token, m_active_key, m_active_ball
    Private m_left, m_confirmed, m_return_seen, m_mechanical_attempt
    Private m_count_token, m_recounting, m_stopped, m_coil_enabled, m_enable_token

    Public Default Function Init(name)
        m_name = "balldevice_" & name
        m_ball_switches = Array()
        m_switch_balls = Array()
        m_switch_due = Array()
        m_eject_events = Array()
        m_eject_all_events = Array()
        m_eject_targets = Array()
        m_player_controlled_eject_event = ""
        m_eject_callback = Null
        Set m_present = CreateObject("Scripting.Dictionary")
        Set m_lost = CreateObject("Scripting.Dictionary")
        Set m_queue = CreateObject("Scripting.Dictionary")
        Set m_active_ball = Nothing
        m_next_request = 0
        m_request = ""
        m_active_key = ""
        m_attempt = 0
        m_token = 0
        m_count_token = 0
        m_enable_token = 0
        m_eject_timeout = 1000
        m_eject_enable_time = 0
        m_entrance_count_delay = 500
        m_exit_count_delay = 500
        m_retry_delay = 500
        m_max_attempts = 0
        m_missing_timeout = 20000
        m_idle_missing_timeout = 3000
        m_eject_switch = 0
        m_incoming_balls = 0
        m_default_device = False
        m_mechanical_eject = False
        m_auto_fire_on_unexpected_ball = True
        m_exclude_from_ball_search = False
        m_debug = False
        m_left = False
        m_confirmed = False
        m_return_seen = False
        m_mechanical_attempt = False
        m_recounting = False
        m_stopped = False
        m_coil_enabled = False
        m_state = "idle"
        ' GLF queue events are asynchronous. The kick must be the final handler,
        ' after all higher-priority handlers (including wait_for blockers).
        AddPinEventListener m_name & "_ball_eject_attempt", m_name & "_perform_eject", _
            "BallDeviceEventHandler", -2147483647, Array("perform_eject", Me)
        glf_ball_devices.Add name, Me
        Set Init = Me
    End Function

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Balls(): Balls = m_present.Count: End Property
    Public Property Get HasBall(): HasBall = (m_present.Count > 0): End Property
    Public Property Get IncomingBalls(): IncomingBalls = m_incoming_balls: End Property
    Public Property Get IsReadyToReceive()
        IsReadyToReceive = False
        If m_stopped Or m_state = "eject_broken" Or m_left Then Exit Property
        If m_state = "waiting_for_queue" Or m_state = "ejecting" Or m_state = "retry_wait" Then Exit Property
        If SettleWait() > 0 Then Exit Property
        Dim raw, occupied : Set raw = RawBalls()
        occupied = raw.Count
        If m_present.Count > occupied Then occupied = m_present.Count
        IsReadyToReceive = (occupied + m_incoming_balls < UBound(m_ball_switches) + 1)
    End Property
    Public Property Get PendingEjects(): PendingEjects = m_queue.Count: End Property
    Public Property Get State(): State = m_state: End Property
    Public Property Get EjectAttempts(): EjectAttempts = m_attempt: End Property
    Public Property Get GetValue(value)
        Select Case LCase(value)
            Case "balls": GetValue = m_present.Count
            Case "state": GetValue = m_state
            Case "incoming_balls": GetValue = m_incoming_balls
            Case "pending_ejects": GetValue = m_queue.Count
        End Select
    End Property

    Public Property Let DefaultDevice(value)
        m_default_device = CBool(value)
        If m_default_device Then Set glf_plunger = Me
    End Property
    Public Property Let AddIncomingBalls(value)
        m_incoming_balls = m_incoming_balls + CLng(value)
        If m_incoming_balls < 0 Then m_incoming_balls = 0
    End Property
    Public Property Let EjectCallback(value): m_eject_callback = value: End Property
    Public Property Let EjectTimeout(value): m_eject_timeout = PositiveMs(value): End Property
    Public Property Let EjectEnableTime(value): m_eject_enable_time = Nonnegative(value): End Property
    Public Property Let EntranceCountDelay(value): m_entrance_count_delay = Nonnegative(value): End Property
    Public Property Let ExitCountDelay(value): m_exit_count_delay = Nonnegative(value): End Property
    Public Property Let EjectRetryDelay(value): m_retry_delay = PositiveMs(value): End Property
    Public Property Let MaxEjectAttempts(value): m_max_attempts = Nonnegative(value): End Property
    Public Property Let BallMissingTimeout(value): m_missing_timeout = PositiveMs(value): End Property
    Public Property Let IdleMissingBallTimeout(value): m_idle_missing_timeout = PositiveMs(value): End Property
    Public Property Let EjectSwitchIndex(value): m_eject_switch = Nonnegative(value): End Property
    Public Property Let MechanicalEject(value)
        If CBool(value) And UBound(m_ball_switches) > 0 Then
            Err.Raise 5, m_name, "MechanicalEject requires one occupancy switch."
        End If
        m_mechanical_eject = CBool(value)
    End Property
    Public Property Let ExcludeFromBallSearch(value): m_exclude_from_ball_search = CBool(value): End Property
    Public Property Let AutoFireOnUnexpectedBall(value): m_auto_fire_on_unexpected_ball = CBool(value): End Property
    Public Property Get AutoFireOnUnexpectedBall(): AutoFireOnUnexpectedBall = m_auto_fire_on_unexpected_ball: End Property
    Public Property Let Debug(value): m_debug = CBool(value): End Property

    Public Property Let BallSwitches(value)
        Dim i, raw : Set raw = RawBalls()
        If raw.Count > 0 Or m_present.Count > 0 Or m_queue.Count > 0 Then
            Err.Raise 5, m_name, "Cannot change BallSwitches while occupied or ejects are pending."
        End If
        If m_mechanical_eject And UBound(value) > 0 Then
            Err.Raise 5, m_name, "MechanicalEject requires one occupancy switch."
        End If
        For i = 0 To UBound(m_ball_switches)
            RemovePinEventListener m_ball_switches(i) & "_active", m_name & "_ball_enter_" & i
            RemovePinEventListener m_ball_switches(i) & "_inactive", m_name & "_ball_exit_" & i
        Next
        m_ball_switches = value
        m_switch_balls = Array()
        m_switch_due = Array()
        If UBound(value) >= 0 Then
            ReDim m_switch_balls(UBound(value))
            ReDim m_switch_due(UBound(value))
        End If
        For i = 0 To UBound(value)
            Set m_switch_balls(i) = CreateObject("Scripting.Dictionary")
            m_switch_due(i) = 0
            AddPinEventListener value(i) & "_active", m_name & "_ball_enter_" & i, _
                "BallDeviceEventHandler", 1000, Array("ball_entering", Me, i)
            AddPinEventListener value(i) & "_inactive", m_name & "_ball_exit_" & i, _
                "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, i)
        Next
    End Property
    Public Property Let EjectEvents(value)
        Dim evt
        For Each evt In m_eject_events
            RemovePinEventListener evt, m_name & "_eject_one"
        Next
        m_eject_events = value
        For Each evt In value
            AddPinEventListener evt, m_name & "_eject_one", "BallDeviceEventHandler", 1000, Array("ball_eject", Me)
        Next
    End Property
    Public Property Let EjectAllEvents(value)
        Dim evt
        For Each evt In m_eject_all_events
            RemovePinEventListener evt, m_name & "_eject_all"
        Next
        m_eject_all_events = value
        For Each evt In value
            AddPinEventListener evt, m_name & "_eject_all", "BallDeviceEventHandler", 1000, Array("ball_eject_all", Me)
        Next
    End Property
    Public Property Let EjectTargets(value)
        If m_request <> "" Then Err.Raise 5, m_name, "Cannot change EjectTargets during an eject."
        m_eject_targets = value
    End Property
    Public Property Let PlayerControlledEjectEvent(value)
        If m_player_controlled_eject_event <> "" Then
            RemovePinEventListener m_player_controlled_eject_event, m_name & "_player_eject"
        End If
        m_player_controlled_eject_event = value
        If value <> "" Then
            AddPinEventListener value, m_name & "_player_eject", "BallDeviceEventHandler", 1000, Array("player_eject", Me)
        End If
    End Property

    ' Seed known balls at table initialization, before normal Hit/UnHit handling.
    ' VPX does not expose a generic occupied-switch -> ball-object lookup.
    Public Sub SeedBall(ball, switch)
        CheckSwitch switch
        If m_request <> "" Then Err.Raise 5, m_name, "Cannot seed during an eject."
        Dim key : key = BallKey(ball)
        If Not m_switch_balls(switch).Exists(key) Then m_switch_balls(switch).Add key, ball
        If Not m_present.Exists(key) Then m_present.Add key, ball
        m_switch_due(switch) = GameTime
    End Sub

    Public Sub BallEntering(ball, switch)
        CheckSwitch switch
        Dim key : key = BallKey(ball)
        If m_switch_balls(switch).Exists(key) Then Exit Sub
        m_switch_balls(switch).Add key, ball
        m_switch_due(switch) = GameTime + m_entrance_count_delay
        If key = m_active_key And m_left Then
            ' Invalidate a confirm as soon as a return is observed, before debounce.
            m_confirmed = False
            m_return_seen = True
            RemoveDelay m_name & "_confirm"
            RemoveDelay m_name & "_missing"
        End If
        If Not m_stopped Then ScheduleRecount
    End Sub
    Public Sub BallEnter(ball, switch)
        ' Compatibility entry point. All arrivals go through the same debounce.
        BallEntering ball, switch
    End Sub
    Public Sub BallExiting(ball, switch)
        CheckSwitch switch
        Dim key : key = BallKey(ball)
        If Not m_switch_balls(switch).Exists(key) Then Exit Sub
        m_switch_balls(switch).Remove key
        m_switch_due(switch) = GameTime + m_exit_count_delay
        If Not m_stopped Then ScheduleRecount
    End Sub

    Private Function RawBalls()
        Dim result, i, key : Set result = CreateObject("Scripting.Dictionary")
        For i = 0 To UBound(m_switch_balls)
            For Each key In m_switch_balls(i).Keys
                If Not result.Exists(key) Then result.Add key, m_switch_balls(i)(key)
            Next
        Next
        Set RawBalls = result
    End Function
    Private Function SettleWait()
        Dim i, remaining : remaining = 0
        For i = 0 To UBound(m_switch_due)
            If m_switch_due(i) - GameTime > remaining Then remaining = m_switch_due(i) - GameTime
        Next
        SettleWait = remaining
    End Function
    Private Sub ScheduleRecount()
        m_count_token = m_count_token + 1
        Schedule "recount", m_count_token, SettleWait(), "_recount"
    End Sub
    Public Sub Recount()
        If m_stopped Or m_recounting Then Exit Sub
        If SettleWait() > 0 Then
            ScheduleRecount
            Exit Sub
        End If
        m_recounting = True
        Dim current, previous, key, ball, old_count
        Set current = RawBalls()
        Set previous = m_present
        old_count = previous.Count
        Set m_present = current
        ' Reconcile identities, not the number of switch edges. Switch-to-switch
        ' movement and overlapping switches do not constitute a departure/arrival.
        For Each key In previous.Keys
            If Not current.Exists(key) Then
                Set ball = previous(key)
                If key = m_active_key And (m_state = "ejecting" Or m_state = "waiting_for_confirm") Then
                    Departed
                ElseIf m_mechanical_eject And (m_state = "idle" Or m_state = "waiting_for_ball" Or _
                        m_state = "waiting_for_position" Or m_state = "waiting_for_player" Or m_state = "waiting_for_queue") Then
                    BeginMechanical ball
                Else
                    If Not m_lost.Exists(key) Then m_lost.Add key, GameTime + m_idle_missing_timeout
                    ScheduleLostCheck
                    If key = m_active_key And m_state = "waiting_for_queue" Then
                        Dim lost_kwargs : Set lost_kwargs = GlfKwargs()
                        lost_kwargs.Add "balls", 1
                        lost_kwargs.Add "retry", False
                        lost_kwargs.Add "num_attempts", m_attempt
                        lost_kwargs.Add "reason", "ball_disappeared_before_eject"
                        DispatchPinEvent m_name & "_ball_eject_failed", lost_kwargs
                        FinishRequest
                    End If
                    CancelBoundRequests key
                    DispatchPinEvent m_name & "_ball_exiting", Null
                End If
            End If
        Next
        For Each key In current.Keys
            If Not previous.Exists(key) Then
                If key = m_active_key And m_left Then
                    ' This is the ejected ball returning, not a new capture.
                ElseIf m_lost.Exists(key) Then
                    m_lost.Remove key
                    Log "Recovered idle ball " & key
                Else
                    Received current(key)
                End If
            End If
        Next
        If m_left Then
            If current.Exists(m_active_key) Then
                FailAttempt "ball_returned"
            ElseIf m_return_seen Then
                ' A return bounced out again before settling. Wait for a fresh confirm.
                m_return_seen = False
                Schedule "confirm_timeout", m_token, m_eject_timeout, "_confirm"
            ElseIf m_confirmed Then
                CompleteEject
            End If
        End If
        If old_count <> current.Count Then
            Dim kwargs : Set kwargs = GlfKwargs()
            kwargs.Add "balls", current.Count
            DispatchPinEvent m_name & "_ball_count_changed", kwargs
        End If
        m_recounting = False
        SchedulePump 1
    End Sub

    Private Sub Received(ball)
        Dim unclaimed, key : key = BallKey(ball)
        unclaimed = 1
        If m_incoming_balls > 0 Then
            m_incoming_balls = m_incoming_balls - 1
            unclaimed = 0
        End If
        ' Preserve GLF's numeric relay contract and expected-arrival behaviour.
        If unclaimed > 0 Then unclaimed = DispatchRelayPinEvent(m_name & "_ball_enter", 1)
        DispatchPinEvent m_name & "_ball_entered", Null
        ' A request queued while empty owns this arrival. Do not add an automatic
        ' request for the same ball on top of that explicit request.
        If Not ReservePendingRequest(key) Then
            If unclaimed > 0 And m_auto_fire_on_unexpected_ball Then Enqueue key, 500
        End If
    End Sub

    Public Sub Eject()
        If m_stopped Or m_state = "eject_broken" Then Exit Sub
        If m_state = "waiting_for_player" Then
            ' Upgrade the existing returned mechanical ball to a powered eject.
            m_mechanical_attempt = False
            m_state = "waiting_for_position"
            SchedulePump 1
        Else
            Enqueue "", 0
        End If
    End Sub
    Public Sub PlayerEject()
        ' A launch button pressed while empty must not arm a future ball.
        If Not HasBall Then Exit Sub
        If m_state = "waiting_for_player" Then
            Eject
        ElseIf m_queue.Count = 0 Then
            Eject
        End If
    End Sub
    Public Sub EjectBalls(balls)
        Dim i, count : count = Nonnegative(balls)
        For i = 1 To count
            Eject
        Next
    End Sub
    Public Sub EjectAll()
        Dim count, reserved : reserved = m_queue.Count
        ' A departed ball is still queued until confirmed, but no longer counted.
        If m_left Then reserved = reserved - 1
        count = m_present.Count - reserved
        If count > 0 Then EjectBalls count
    End Sub
    Private Sub Enqueue(key, delay_ms)
        If m_stopped Or m_state = "eject_broken" Then Exit Sub
        ' Bind known balls once, so automatic and explicit requests cannot both
        ' reserve the same ball. Requests exceeding supply remain unbound.
        If key = "" Then key = FindUnreservedBall()
        m_next_request = m_next_request + 1
        m_queue.Add CStr(m_next_request), Array(key, GameTime + delay_ms)
        SchedulePump 1
    End Sub
    Private Function BallReserved(key)
        Dim id, data : BallReserved = False
        For Each id In m_queue.Keys
            data = m_queue(id)
            If data(0) = key Then
                BallReserved = True
                Exit Function
            End If
        Next
    End Function
    Private Function FindUnreservedBall()
        Dim key : FindUnreservedBall = ""
        If m_eject_switch <= UBound(m_switch_balls) Then
            For Each key In m_switch_balls(m_eject_switch).Keys
                If m_present.Exists(key) Then
                    If Not BallReserved(key) Then
                        FindUnreservedBall = key
                        Exit Function
                    End If
                End If
            Next
        End If
        For Each key In m_present.Keys
            If Not BallReserved(key) Then
                FindUnreservedBall = key
                Exit Function
            End If
        Next
    End Function
    Private Function ReservePendingRequest(key)
        Dim id, data
        ReservePendingRequest = BallReserved(key)
        If ReservePendingRequest Then Exit Function
        For Each id In m_queue.Keys
            data = m_queue(id)
            If data(0) = "" Then
                m_queue(id) = Array(key, data(1))
                ReservePendingRequest = True
                Exit Function
            End If
        Next
    End Function
    Private Sub SchedulePump(delay_ms)
        If m_stopped Or m_state = "eject_broken" Then Exit Sub
        If m_queue.Count = 0 Then Exit Sub
        Schedule "pump", m_token, delay_ms, "_pump"
    End Sub
    Public Sub Pump()
        If m_stopped Or m_recounting Or m_state = "eject_broken" Then Exit Sub
        Select Case m_state
            Case "waiting_for_queue", "ejecting", "waiting_for_confirm", "retry_wait", "waiting_for_player"
                Exit Sub
        End Select
        If m_queue.Count = 0 Then
            m_state = "idle"
            Exit Sub
        End If
        If SettleWait() > 0 Then
            ScheduleRecount
            Exit Sub
        End If
        If m_eject_switch > UBound(m_ball_switches) Then
            Err.Raise 5, m_name, "EjectSwitchIndex must identify a configured BallSwitch."
        End If
        If IsNull(m_eject_callback) Then
            m_state = "waiting_for_player"
            Exit Sub
        End If
        Dim keys, request_data, key, candidate
        If m_request = "" Then
            keys = m_queue.Keys
            m_request = keys(0)
            m_attempt = 0
        End If
        request_data = m_queue(m_request)
        If request_data(1) > GameTime Then
            SchedulePump request_data(1) - GameTime
            Exit Sub
        End If
        candidate = ""
        For Each key In m_switch_balls(m_eject_switch).Keys
            If m_present.Exists(key) Then
                If request_data(0) = "" Or request_data(0) = key Then
                    candidate = key
                    Exit For
                End If
            End If
        Next
        If candidate = "" Then
            If m_present.Count = 0 Then
                m_state = "waiting_for_ball"
            Else
                m_state = "waiting_for_position"
            End If
            Exit Sub
        End If
        m_active_key = candidate
        m_queue(m_request) = Array(candidate, request_data(1))
        Set m_active_ball = m_present(candidate)
        m_token = m_token + 1
        m_mechanical_attempt = False
        m_left = False
        m_confirmed = False
        m_return_seen = False
        m_state = "waiting_for_queue"
        Dim kwargs : Set kwargs = GlfKwargs()
        kwargs.Add "glf_eject_token", m_token
        kwargs.Add "source", Me
        kwargs.Add "balls", 1
        kwargs.Add "num_attempts", m_attempt
        DispatchQueuePinEvent m_name & "_ball_eject_attempt", kwargs
    End Sub
    Public Sub PerformEject(token)
        If m_stopped Or token <> m_token Or m_state <> "waiting_for_queue" Then Exit Sub
        ' The ball may have moved while an external queue handler was blocking.
        If SettleWait() > 0 Then
            Schedule "perform_eject", token, SettleWait(), "_perform"
            Exit Sub
        End If
        If Not m_switch_balls(m_eject_switch).Exists(m_active_key) Then
            m_state = "waiting_for_position"
            m_active_key = ""
            Set m_active_ball = Nothing
            Recount
            Exit Sub
        End If
        m_attempt = m_attempt + 1
        m_state = "ejecting"
        m_left = False
        m_confirmed = False
        m_return_seen = False
        RegisterConfirmListeners
        Schedule "departure_timeout", m_token, m_eject_timeout, "_departure"
        DispatchPinEvent m_name & "_ejecting_ball", Null
        Log "Eject attempt " & m_attempt & ", ball " & m_active_key
        FireCallback m_active_ball
    End Sub
    Private Sub Departed()
        If m_left Then Exit Sub
        m_left = True
        m_state = "waiting_for_confirm"
        RemoveDelay m_name & "_departure"
        Schedule "confirm_timeout", m_token, m_eject_timeout, "_confirm"
        DispatchPinEvent m_name & "_ball_exiting", Null
    End Sub
    Private Sub BeginMechanical(ball)
        If m_request = "" Then
            If m_queue.Count = 0 Then Enqueue BallKey(ball), 0
            Dim keys : keys = m_queue.Keys
            m_request = keys(0)
            m_attempt = 0
        End If
        m_token = m_token + 1
        m_attempt = m_attempt + 1
        m_active_key = BallKey(ball)
        Set m_active_ball = ball
        m_mechanical_attempt = True
        m_left = False
        m_confirmed = False
        m_return_seen = False
        RegisterConfirmListeners
        DispatchPinEvent m_name & "_ejecting_ball", Null
        Departed
    End Sub
    Private Sub RegisterConfirmListeners()
        Dim evt
        For Each evt In m_eject_targets
            RemovePinEventListener evt & "_active", m_name & "_confirm"
            AddPinEventListener evt & "_active", m_name & "_confirm", "BallDeviceEventHandler", _
                1000, Array("confirm", Me, m_token)
        Next
    End Sub
    Public Sub ConfirmEject(ball, token)
        If m_stopped Or token <> m_token Then Exit Sub
        If m_state <> "ejecting" And m_state <> "waiting_for_confirm" Then Exit Sub
        If Not IsObject(ball) Then Exit Sub
        If BallKey(ball) <> m_active_key Then Exit Sub
        Dim raw : Set raw = RawBalls()
        If raw.Exists(m_active_key) Then Exit Sub
        m_confirmed = True
        ' A fast confirmation can precede exit debounce. Latch it until recount.
        If m_left And SettleWait() = 0 Then CompleteEject
    End Sub
    Public Sub BallExitSuccess(ball)
        ' Legacy public confirmation entry point; still validates the active ball.
        ConfirmEject ball, m_token
    End Sub
    Private Sub CompleteEject()
        If m_request = "" Or Not m_left Then Exit Sub
        Dim raw : Set raw = RawBalls()
        If raw.Exists(m_active_key) Or SettleWait() > 0 Then Exit Sub
        Log "Eject succeeded, ball " & m_active_key
        FinishRequest
        DispatchPinEvent m_name & "_ball_eject_success", Null
    End Sub
    Private Sub FailAttempt(reason)
        If m_request = "" Then Exit Sub
        CancelAttempt
        Dim retry, kwargs : retry = True
        If m_max_attempts > 0 Then retry = (m_attempt < m_max_attempts)
        Set kwargs = GlfKwargs()
        kwargs.Add "balls", 1
        kwargs.Add "retry", retry
        kwargs.Add "num_attempts", m_attempt
        kwargs.Add "reason", reason
        DispatchPinEvent m_name & "_ball_eject_failed", kwargs
        m_left = False
        m_confirmed = False
        m_return_seen = False
        If Not retry Then
            m_state = "eject_broken"
            DispatchPinEvent m_name & "_broken", Null
        ElseIf m_mechanical_attempt Then
            m_state = "waiting_for_player"
        Else
            m_state = "retry_wait"
            Schedule "retry", m_token, m_retry_delay, "_retry"
        End If
        Log "Eject failed: " & reason
    End Sub
    Private Sub FinishRequest()
        CancelAttempt
        If m_queue.Exists(m_request) Then m_queue.Remove m_request
        m_request = ""
        m_active_key = ""
        Set m_active_ball = Nothing
        m_left = False
        m_confirmed = False
        m_return_seen = False
        m_mechanical_attempt = False
        m_state = "idle"
        SchedulePump 1
    End Sub
    Private Sub CancelAttempt()
        Dim evt
        m_token = m_token + 1
        For Each evt In Array("_departure", "_confirm", "_missing", "_retry", "_perform", _
                "_departure_timeout", "_confirm_timeout", "_missing_timeout")
            RemoveDelay m_name & evt
        Next
        For Each evt In m_eject_targets
            RemovePinEventListener evt & "_active", m_name & "_confirm"
        Next
        EjectEnableComplete
    End Sub
    Private Sub CancelBoundRequests(key)
        Dim id, data
        For Each id In m_queue.Keys
            data = m_queue(id)
            If data(0) = key Then
                If id = m_request Then
                    FinishRequest
                Else
                    m_queue.Remove id
                End If
            End If
        Next
    End Sub

    Public Sub OnTimer(kind, token)
        If m_stopped Then Exit Sub
        If kind = "recount" Then
            If token = m_count_token Then Recount
            Exit Sub
        End If
        If kind = "lost" Then
            ClearLostBalls
            Exit Sub
        End If
        If kind = "enable_complete" Then
            If token = m_enable_token Then EjectEnableComplete
            Exit Sub
        End If
        If kind = "pump" Then
            ' Pump has no captured request and only examines the current queue.
            Pump
            Exit Sub
        End If
        If token <> m_token Then Exit Sub
        Select Case kind
            Case "perform_eject": PerformEject token
            Case "retry"
                If m_state = "retry_wait" Then
                    m_state = "waiting_for_position"
                    Pump
                End If
            Case "departure_timeout", "confirm_timeout", "missing_timeout"
                If SettleWait() > 0 Then
                    Schedule kind, token, SettleWait(), "_" & kind
                    Exit Sub
                End If
                Recount
                If token <> m_token Then Exit Sub
                If kind = "departure_timeout" Then
                    If m_state = "ejecting" Then FailAttempt "ball_did_not_leave"
                ElseIf m_state = "waiting_for_confirm" Then
                    If m_present.Exists(m_active_key) Then
                        FailAttempt "ball_returned"
                    ElseIf m_confirmed Or UBound(m_eject_targets) < 0 Then
                        CompleteEject
                    ElseIf kind = "confirm_timeout" Then
                        Schedule "missing_timeout", m_token, m_missing_timeout, "_missing"
                    Else
                        ' The ball left but never reached the confirmation switch.
                        ' Retire this request; kicking a different ball would over-serve.
                        Dim kwargs : Set kwargs = GlfKwargs()
                        kwargs.Add "balls", 1
                        kwargs.Add "retry", False
                        kwargs.Add "num_attempts", m_attempt
                        kwargs.Add "reason", "confirmation_missing"
                        DispatchPinEvent m_name & "_ball_eject_failed", kwargs
                        ReportMissing m_active_key
                        FinishRequest
                    End If
                End If
        End Select
    End Sub
    Private Sub Schedule(kind, token, delay_ms, suffix)
        If delay_ms < 1 Then delay_ms = 1
        SetDelay m_name & suffix, "BallDeviceEventHandler", Array(Array(kind, Me, token), Null), delay_ms
    End Sub
    Private Sub ScheduleLostCheck()
        Dim key, wait_ms : wait_ms = m_idle_missing_timeout
        For Each key In m_lost.Keys
            If m_lost(key) - GameTime < wait_ms Then wait_ms = m_lost(key) - GameTime
        Next
        If m_lost.Count > 0 Then Schedule "lost", 0, wait_ms, "_lost"
    End Sub
    Public Sub ClearLostBalls()
        Dim key, raw : Set raw = RawBalls()
        For Each key In m_lost.Keys
            If m_lost(key) <= GameTime Then
                If raw.Exists(key) Then
                    ' A return may already be visible but still inside debounce.
                    m_lost(key) = GameTime + SettleWait() + 1
                Else
                    ReportMissing key
                    m_lost.Remove key
                End If
            End If
        Next
        ScheduleLostCheck
    End Sub
    Private Sub ReportMissing(key)
        Dim kwargs : Set kwargs = GlfKwargs()
        kwargs.Add "balls", 1
        kwargs.Add "name", m_name
        kwargs.Add "ball_id", key
        DispatchPinEvent m_name & "_ball_missing", kwargs
        DispatchPinEvent "balldevice_ball_missing", kwargs
        Log "Ball missing: " & key
    End Sub

    Private Sub FireCallback(ball)
        EjectEnableComplete
        If m_eject_enable_time > 0 Then
            m_coil_enabled = True
            m_enable_token = m_enable_token + 1
            Schedule "enable_complete", m_enable_token, m_eject_enable_time, "_enable"
        End If
        GetRef(m_eject_callback)(ball)
    End Sub
    Public Sub EjectEnableComplete()
        RemoveDelay m_name & "_enable"
        m_enable_token = m_enable_token + 1
        If Not m_coil_enabled Then Exit Sub
        m_coil_enabled = False
        If Not IsNull(m_eject_callback) Then GetRef(m_eject_callback)(Null)
    End Sub
    Public Sub BallSearch(phase)
        If m_stopped Or m_exclude_from_ball_search Or m_default_device Then Exit Sub
        If m_state <> "idle" Or m_queue.Count > 0 Then Exit Sub
        If IsNull(m_eject_callback) Then Exit Sub
        If SettleWait() > 0 Then Exit Sub
        If m_present.Count > 0 Then
            If phase > 1 Then Eject
        Else
            FireCallback Null
        End If
    End Sub
    Public Sub StopDevice()
        m_stopped = True
        CancelAttempt
        RemoveDelay m_name & "_pump"
        RemoveDelay m_name & "_recount"
        RemoveDelay m_name & "_lost"
        m_count_token = m_count_token + 1
        m_queue.RemoveAll
        m_request = ""
        m_active_key = ""
        Set m_active_ball = Nothing
        m_left = False
        m_state = "stopped"
    End Sub
    Public Sub Reset()
        StopDevice
        m_stopped = False
        m_state = "idle"
        m_attempt = 0
        ScheduleRecount
        ScheduleLostCheck
    End Sub
    Public Function ToYaml()
        ToYaml = "  " & m_name & ":" & vbCrLf & _
            "    ball_switches: " & Join(m_ball_switches, ",") & vbCrLf & _
            "    mechanical_eject: " & LCase(CStr(m_mechanical_eject)) & vbCrLf
    End Function
    Private Function BallKey(ball)
        If Not IsObject(ball) Then Err.Raise 5, m_name, "Switch events must supply a VPX ball object."
        BallKey = CStr(ball.ID)
    End Function
    Private Sub CheckSwitch(index)
        If index < 0 Or index > UBound(m_ball_switches) Then Err.Raise 9, m_name, "Invalid ball switch index."
    End Sub
    Private Function Nonnegative(value)
        If Not IsNumeric(value) Then Err.Raise 5, m_name, "Expected a nonnegative integer."
        If value < 0 Or CDbl(value) <> Fix(CDbl(value)) Then Err.Raise 5, m_name, "Expected a nonnegative integer."
        Nonnegative = CLng(value)
    End Function
    Private Function PositiveMs(value)
        PositiveMs = Nonnegative(value)
        If PositiveMs = 0 Then Err.Raise 5, m_name, "Timeouts must be greater than zero."
    End Function
    Private Sub Log(message)
        If m_debug Then glf_debugLog.WriteToLog m_name, message
    End Sub
End Class

Function BallDeviceEventHandler(args)
    Dim ownProps, evt, device, kwargs
    ownProps = args(0)
    evt = ownProps(0)
    Set device = ownProps(1)
    Select Case evt
        Case "ball_entering", "ball_enter": device.BallEntering args(1), ownProps(2)
        Case "ball_exiting": device.BallExiting args(1), ownProps(2)
        Case "ball_eject": device.Eject
        Case "ball_eject_all": device.EjectAll
        Case "player_eject": device.PlayerEject
        Case "confirm": device.ConfirmEject args(1), ownProps(2)
        Case "perform_eject"
            If IsObject(args(1)) Then
                Set kwargs = args(1)
                If kwargs.Exists("glf_eject_token") Then device.PerformEject kwargs("glf_eject_token")
                If kwargs.Exists("wait_for") Then kwargs.Remove "wait_for"
            Else
                device.OnTimer evt, ownProps(2)
            End If
        Case Else: device.OnTimer evt, ownProps(2)
    End Select
    ' GLF requires every queue-event handler to return its kwargs object.
    If IsObject(args(1)) Then
        Set BallDeviceEventHandler = args(1)
    Else
        BallDeviceEventHandler = args(1)
    End If
End Function
