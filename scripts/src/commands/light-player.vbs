
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
        Dim yaml
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
        Exit Function
        'glf_debugLog.WriteToLog "LightPlayer", "Removing Light Seq" & mode & "_" & key
    Else
        If UBound(lights) = -1 Then
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
    End If
    Set LightPlayerCallbackHandler = shows_added
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
