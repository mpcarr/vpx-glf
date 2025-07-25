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
