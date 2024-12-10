
Class GlfEvent
	Private m_raw, m_name, m_event, m_condition, m_delay
  
    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventName() : EventName = m_event : End Property
    Public Property Get Condition() : Condition = m_condition : End Property
    Public Property Get Delay() : Delay = m_delay : End Property
    Public Property Get Raw() : Raw = m_raw : End Property

    Public Function Evaluate()
        If Not IsNull(m_condition) Then
            Evaluate = GetRef(m_condition)()
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
        m_delay = parsedEvent(3)
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

    Public Property Let FallbackEvent(value) : m_fallback_event = value : End Property
    Public Property Let ForceAll(value) : m_force_all = value : End Property
    Public Property Let ForceDifferent(value) : m_force_different = value : End Property
    Public Property Let DisableRandom(value) : m_disable_random = value : End Property

    Public Function Evaluate()
        If Not IsNull(m_condition) Then
            Evaluate = GetRef(m_condition)()
        Else
            Evaluate = True
        End If
    End Function

	Public default Function init(evt, mode, key)
        m_parent_key = evt
        m_key = key
        m_mode = mode
        m_fallback_event = Empty
        m_force_all = True
        m_force_different = True
        m_disable_random = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_weights = CreateObject("Scripting.Dictionary")
        Set m_eventIndexMap = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

    Public Sub Add(evt, weight)
        Dim newEvent : Set newEvent = (new GlfEvent)(evt)
        m_events.Add newEvent.Raw, newEvent
        m_weights.Add newEvent.Raw, weight
        m_eventIndexMap.Add newEvent.Raw, UBound(m_events.Keys)
    End Sub

    Public Function GetNextRandomEvent()

        Dim valid_events, event_to_fire
        Dim event_keys, event_items
        Dim i, count
        
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
        Dim randomIdx
        If m_disable_random = False Then
            Randomize
            randomIdx = Int(Rnd() * (UBound(valid_events.Keys)-LBound(valid_events.Keys) + 1) + LBound(valid_events.Keys))
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
            randomIdx = 0
        End If
        
        SetPlayerState "random_" & m_mode & "_" & m_key & "_last", valid_events.Items()(randomIdx).Raw
        SetPlayerState "random_" & m_mode & "_" & m_key & "_" & valid_events.Items()(randomIdx).Raw, True

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

        GetNextRandomEvent = valid_events.Items()(randomIdx).EventName

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

End Class