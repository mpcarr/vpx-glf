
Class GlfPushTransition

    Private m_name
    Private m_output_length
    Private m_config
    Private m_collapse_dots
    Private m_collapse_commas
    Private m_use_dots_for_commas
    Private m_current_step
    Private m_total_steps

    Private m_direction
    Private m_text
    Private m_text_color

    ' Properties
    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(value): m_name = value: End Property

    Public Property Let OutputLength(value): m_output_length = value: End Property
    Public Property Get OutputLength(): OutputLength = m_output_length: End Property

    Public Property Let Config(value): m_config = value: End Property
    Public Property Get Config(): Config = m_config: End Property

    Public Property Let CollapseDots(value): m_collapse_dots = value: End Property
    Public Property Get CollapseDots(): CollapseDots = m_collapse_dots: End Property

    Public Property Let CollapseCommas(value): m_collapse_commas = value: End Property
    Public Property Get CollapseCommas(): CollapseCommas = m_collapse_commas: End Property

    Public Property Let UseDotsForCommas(value): m_use_dots_for_commas = value: End Property
    Public Property Get UseDotsForCommas(): UseDotsForCommas = m_use_dots_for_commas: End Property

    Public Property Let CurrentStep(value): m_current_step = value: End Property
    Public Property Get CurrentStep(): CurrentStep = m_current_step: End Property

    Public Property Let TotalSteps(value): m_total_steps = value: End Property
    Public Property Get TotalSteps(): TotalSteps = m_total_steps: End Property

    ' Initialize the class
    Public Function Init(name, output_length, collapse_dots, collapse_commas, use_dots_for_commas, config)
        m_name = "transition_" & name
        m_output_length = output_length
        m_collapse_dots = collapse_dots
        m_collapse_commas = collapse_commas
        m_use_dots_for_commas = use_dots_for_commas
        m_config = config
        m_current_step = 0
        m_total_steps = 0

        m_direction = "right"
        m_text = Empty
        m_color = Empty
        Set Init = Me
    End Function

    Public Function GetStepCount()
        GetStepCount = m_output_length + Len(m_text)
    End Function

    Public Function GetTransitionStep(current_step, current_text, new_text, current_colors, new_colors)

        If current_step < 0 or current_step >= GetStepCount() Then
            Log "Step id out of range"
            Exit Function
        End If
        
        Dim current_display_text
        Dim current_char_list : current_char_list = Glf_SegmentTextCreateCharacters(current_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, current_colors)
        Set current_display_text = (new GlfSegmentDisplayText)(current_char_list, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas)
        
        Dim new_display_text    
        Dim new_char_list : new_char_list = Glf_SegmentTextCreateCharacters(new_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, new_colors)
        Set new_display_text = (new GlfSegmentDisplayText)(new_char_list, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas) 

        Dim text_color,transition_text
        If Not IsEmpty(m_text) Then
            If IsArray(new_colors) and IsEmpty(m_text_color) Then
                text_color = Array(new_colors(0))
            Else
                text_color = m_text_color
            End If
            Dim trnasition_char_list : trnasition_char_list = Glf_SegmentTextCreateCharacters(m_text, Len(m_text), m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, text_color)
            Set transition_text = (new GlfSegmentDisplayText)(trnasition_char_list,m_collapse_dots, m_collapse_commas, m_use_dots_for_commas) 
        Else
            Set transition_text = (new GlfSegmentDisplayText)(Array(),m_collapse_dots, m_collapse_commas, m_use_dots_for_commas) 
        End If
        Dim start_idx, end_idx
        If m_direction = "right" Then
            Dim temp_list : Set temp_list = new_display_text
            temp_list.Extend transition_text.Text()
            temp_list.Extend current_display_text.Text()

            
            start_idx = m_output_length + Len(m_text) - (current_step + 1)
            end_idx = 2 * m_output_length + Len(m_text) - (current_step + 1) - 1
            GetTransitionStep = SliceArray(temp_list, start_idx, end_idx)
        End If

        If m_direction = "left" Then
            temp_list = current_display_text
            temp_list.Extend transition_text.Text()
            temp_list.Extend new_display_text.Text()

            
            start_idx = current_step + 1
            end_idx = current_step + 1 + m_output_length - 1
            GetTransitionStep = SliceArray(temp_list, start_idx, end_idx)

        End If

    End Function

    ' Start the transition process
    Public Sub StartTransition(current_text, new_text)
        m_current_step = 0
        m_total_steps = GetStepCount()
        If m_total_steps = 0 Then
            m_total_steps = 1 ' Default to at least one step
        End If

        While m_current_step < m_total_steps
            Dim result
            result = GetTransitionStep(m_current_step, current_text, new_text)
            m_base_device.Log "Step " & m_current_step & ": " & result
            m_current_step = m_current_step + 1
        Wend
        m_base_device.Log "Transition complete for: " & m_name
    End Sub

End Class
