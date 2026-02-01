
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
        If IsNull(duration) Then
            new_step.Raw = "time"
        Else
            new_step.Raw = "duration"
        End If

        'Add a empty first step if if show does not start right away
        If UBound(m_steps.Keys) = -1 Then
            If Not IsNull(new_step.Time) And new_step.Time <> 0 Then
                Dim empty_step : Set empty_step = (new GlfShowStep)()
                empty_step.Duration = new_step.Time
                'm_total_step_time = new_step.Time
                m_steps.Add CStr(UBound(m_steps.Keys())+1), empty_step        
            End If
        End If
        

        glf_debugLog.WriteToLog "Show:" & m_name, "Adding Step. Abs: " & absolute_time & ". Rel: " & relative_time & ". Duration: " & duration

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
                        prevStep.Duration = Round(new_step.Time - m_total_step_time, 2)
                        glf_debugLog.WriteToLog "Show:" & m_name, new_step.Time
                        glf_debugLog.WriteToLog "Show:" & m_name, m_total_step_time
                        glf_debugLog.WriteToLog "Show:" & m_name, prevStep.Duration
                    End If
                Else
                    glf_debugLog.WriteToLog "Show:" & m_name, "Setting Duration to 1"
                    prevStep.Duration = 1
                End If
            End If
            m_total_step_time = m_total_step_time + prevStep.Duration
        Else
            If IsNull(new_step.Duration) Then
                m_total_step_time = m_total_step_time + new_step.Time
            Else
                m_total_step_time = m_total_step_time + new_step.Duration
            End If
        End If
        glf_debugLog.WriteToLog "Show:" & m_name, "TOTAL STEP TIME: " & m_total_step_time
        m_steps.Add CStr(UBound(m_steps.Keys())+1), new_step
        Set AddStep = new_step
    End Function

    Public Function ToYaml()
        Dim yaml, show_step, x
        x=0
        For Each show_step in m_steps.Items()
            If Not IsNull(show_step.AbsoluteTime) Then
                If x = 0 Then
                    yaml = yaml & "- time: " & show_step.AbsoluteTime & vbCrLf
                Else
                    yaml = yaml & "- time: " & show_step.AbsoluteTime & "s" & vbCrLf
                End If
                If Not IsNull(show_step.Duration) And show_step.Raw = "duration" Then
                    yaml = yaml & "  duration: " & show_step.Duration & "s" & vbCrLf
                End If
            Else
                If Not IsNull(show_step.Duration) And show_step.Raw = "duration" Then
                    yaml = yaml & "- duration: " & show_step.Duration & "s" & vbCrLf
                Else  
                    If Not IsNull(show_step.RelativeTime) Then
                        yaml = yaml & "- time: +" & show_step.RelativeTime & "s" & vbCrLf
                    End If
                End If
            End If
            yaml = yaml & show_step.toYaml() & vbCrLf
            x=x+1
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
    Private m_mode

    Public Property Get CacheName(): CacheName = m_show_name & "_" & m_internal_cache_id & "_" & ShowSettings.InternalCacheId: End Property
    Public Property Get Tokens(): Set Tokens = m_tokens : End Property

    Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property        
    
    Public Property Get Mode(): Mode = m_mode End Property
    Public Property Let Mode(input): m_mode = input End Property

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
    
    Public default Function init(rname, rkey, show_settings, priority, tokens, cache_id, mode)
        m_show_name = rname
        m_key = rkey
        m_current_step = 0
        m_priority = priority
        m_mode = mode
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
                Set new_running_show = (new GlfRunningShow)(show_item.Key, show_item.Key, show_item, running_show.Priority + running_show.ShowSettings.Priority, Null, Null, running_show.Mode)
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
            If useBcp = True Then
                bcpController.PlaySlide slide_item.Slide, running_show.Mode, "", slide_item.Action, slide_item.Expire, running_show.Priority + running_show.ShowSettings.Priority
            End If
        Next
    End If
    If UBound(nextStep.WidgetsInStep().Keys())>-1 Then
        Dim widget_item
        Dim widget_items : widget_items = nextStep.WidgetsInStep().Items()
        For Each widget_item in widget_items
            If useBcp = True Then
                bcpController.PlayWidget widget_item.Widget, running_show.Mode, "", running_show.Priority + running_show.ShowSettings.Priority, widget_item.Expire
            End If
        Next
    End If
    If UBound(nextStep.SoundsInStep().Keys())>-1 Then
        Dim sound_item
        Dim sound_items : sound_items = nextStep.SoundsInStep().Items()
        For Each sound_item in sound_items
            sound_item.Mode = running_show.Mode
            If sound_item.Action = "stop" Then
                glf_sound_buses(sound_item.Sound.Bus).StopSoundWithKey sound_item.Sound.File
            Else
                glf_sound_buses(sound_item.Sound.Bus).Play sound_item
            End If
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
        If running_show.Loops = -1 Or running_show.Loops > 0 Then
            If running_show.Loops > 0 Then
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

    Private m_lights, m_shows, m_dofs, m_slides, m_widgets, m_sounds, m_time, m_duration, m_isLastStep, m_absTime, m_relTime, m_raw

    Public Property Get Lights(): Lights = m_lights: End Property
    Public Property Let Lights(input) : m_lights = input: End Property
    Public Property Get Raw(): Raw = m_raw: End Property
    Public Property Let Raw(input) : m_raw = input: End Property

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

    Public Property Get WidgetsInStep(): Set WidgetsInStep = m_widgets: End Property
    Public Property Get Widgets(widget)
        Dim new_widget : Set new_widget = (new GlfWidgetPlayerItem)()
        new_widget.Widget = widget
        m_widgets.Add widget & CStr(UBound(m_widgets.Keys)), new_widget
        Set Widgets = new_widget
    End Property

    Public Property Get SoundsInStep(): Set SoundsInStep = m_sounds: End Property
    Public Property Get Sounds(sound)
        Dim new_sound : Set new_sound = (new GlfSoundPlayerItem)(Empty)
        new_sound.Sound = sound
        m_sounds.Add sound & CStr(UBound(m_sounds.Keys)), new_sound
        Set Sounds = new_sound
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
        Set m_widgets = CreateObject("Scripting.Dictionary")
        Set m_sounds = CreateObject("Scripting.Dictionary")
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
                    If light_parts(2) = "stop" Then
                        yaml = yaml & "    " & light_parts(0) & ": stop" & vbCrLf    
                    Else
                        yaml = yaml & "    " & light_parts(0) & ": " & light_parts(2) & "%" & light_parts(1) & vbCrLf
                    End If
                    
                End If
            Next
        End If
        If UBound(m_shows.Keys()) > -1 Then
            yaml = yaml & "  shows:" & vbCrLf
            Dim show
            For Each show in m_shows.Items()
                yaml = yaml & show.ToYaml()
            Next
        End If
        If UBound(m_slides.Keys()) > -1 Then
            yaml = yaml & "  slides:" & vbCrLf
            Dim slide
            For Each slide in m_slides.Items()
                yaml = yaml & slide.ToYaml()
            Next
        End If
        If UBound(m_widgets.Keys()) > -1 Then
            yaml = yaml & "  widgets:" & vbCrLf
            Dim widget
            For Each widget in m_widgets.Items()
                yaml = yaml & widget.ToYaml()
            Next
        End If
        If UBound(m_sounds.Keys()) > -1 Then
            yaml = yaml & "  sounds:" & vbCrLf
            Dim sound
            For Each sound in m_sounds.Items()
                yaml = yaml & sound.ToYaml()
            Next
        End If
        ToYaml = yaml
    End Function

End Class