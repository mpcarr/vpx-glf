
Function CreateGlfSoundBus(name)
	Dim bus : Set bus = (new GlfSoundBus)(name)
	Set CreateGlfSoundBus = bus
End Function

Class GlfSoundBus

    Private m_name
    Private m_simultaneous_sounds
    Private m_current_sounds
    Private m_sound_key
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
        m_sound_key = 0
        Set m_current_sounds = CreateObject("Scripting.Dictionary")
        glf_sound_buses.Add name, Me
        Set Init = Me
    End Function

    Public Sub Play(sound_settings)
        If (UBound(m_current_sounds.Keys)-1) > m_simultaneous_sounds Then
            'TODO: Queue Sound
        Else
            m_current_sounds.Add m_sound_key, sound_settings
            PlaySound(sound_settings.Sound.File)
            SetDelay m_name & "_stop_sound_" & m_sound_key, "Glf_SoundBusStopSoundHandler", Array(m_sound_key, Me), sound_settings.Sound.Duration
            m_sound_key = m_sound_key + 1
        End If
    End Sub

    Public Sub StopSoundWithKey(sound_key)
        Dim sound_settings : Set sound_settings = m_current_sounds(sound_key)
        Dim evt
        For Each evt in sound_settings.Sound.EventsWhenStopped.Items()
            If evt.Evaluate() Then
                DispatchPinEvent evt.EventName, Null
            End If
        Next
        m_current_sounds.Remove sound_key
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
        m_volume = 0.5
        m_priority = 0
        m_duration = 0
        m_max_queue_time = -1 
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