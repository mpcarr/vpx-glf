Class GlfCoverTransition

    Private m_output_length
    Private m_collapse_dots
    Private m_collapse_commas
    Private m_use_dots_for_commas

    Private m_current_step
    Private m_total_steps
    Private m_direction
    Private m_transition_text
    Private m_transition_color

    Private m_current_text
    Private m_new_text
    Private m_current_colors
    Private m_new_colors

    ' Initialize the transition class
    Public Default Function Init(output_length, collapse_dots, collapse_commas, use_dots_for_commas)
        m_output_length = output_length
        m_collapse_dots = collapse_dots
        m_collapse_commas = collapse_commas
        m_use_dots_for_commas = use_dots_for_commas

        m_current_step = 0
        m_total_steps = 0
        m_direction = "right"
        m_transition_text = "" ' Default empty transition text
        m_transition_color = Empty

        Set Init = Me
    End Function

    ' Set transition direction
    Public Property Let Direction(value)
        If value = "left" Or value = "right" Then
            m_direction = value
        Else
            m_direction = "right" ' Default to right
        End If
    End Property

    ' Set transition text
    Public Property Let TransitionText(value)
        m_transition_text = value
    End Property

    ' Start transition
    Public Function StartTransition(current_text, new_text, current_colors, new_colors)
        ' Store text and colors
        m_current_text = "ABCDEGH"
        m_new_text = Space(m_output_length - Len(new_text)) & new_text
        m_current_colors = current_colors
        m_new_colors = new_colors
        ' Calculate total steps for transition
        m_total_steps = m_output_length + Len(m_transition_text)
        'If m_total_steps > 0 Then m_total_steps = m_total_steps
        ' Reset step counter
        m_current_step = 0
        StartTransition = NextStep()
    End Function

    ' Manually call this to progress the transition
    Public Function NextStep()
        If m_current_step >= m_total_steps Then
            NextStep = Null
            Exit Function
        End If
        ' Get the correct transition text for this step
        Dim result
        result = GetTransitionStep(m_current_step)
        ' Increment step counter
        m_current_step = m_current_step + 1

        ' Return the current frame's text
        NextStep = result
    End Function

    ' Get the text for the current step
    Private Function GetTransitionStep(current_step)
        Dim transition_sequence, start_idx, end_idx

        If m_direction = "right" Then
            ' Right cover transition: new_text + transition_text moves in
            transition_sequence = m_new_text & m_transition_text & m_current_text
            start_idx = Len(transition_sequence) - (current_step + m_output_length)
            end_idx = start_idx + m_output_length - 1
        ElseIf m_direction = "left" Then
            ' Left cover transition: transition_text + new_text moves in
            transition_sequence = m_current_text & m_transition_text & m_new_text
            start_idx = current_step
            end_idx = start_idx + m_output_length - 1
        End If

        ' Ensure valid slice indices
        If start_idx < 0 Then start_idx = 0
        If end_idx > Len(transition_sequence) - 1 Then end_idx = Len(transition_sequence) - 1

        ' Extract the correct frame of text
        Dim sliced_text
        sliced_text = Mid(transition_sequence, start_idx + 1, end_idx - start_idx + 1)
        If m_output_length>m_current_step Then
            If m_direction = "right" Then
                sliced_text = m_new_text & m_transition_text & Right(m_current_text, m_output_length-current_step)
            ElseIf m_direction = "left" Then
                sliced_text = Left(m_current_text, m_output_length - m_current_step) & Left(m_transition_text & m_new_text, current_step)
            End If
        End If
        'MsgBox "transition_text-"&transition_sequence&", current_step=" & current_step & ", start_idx=" & start_idx & ", end_idx=" & end_idx & ", text=>" & sliced_text &"<, text_len=>" & Len(sliced_text) &"<, Total Steps: "&m_total_steps
        
        ' Convert only the final sliced text to segment display characters
        GetTransitionStep = Glf_SegmentTextCreateCharacters(sliced_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, m_new_colors)
    End Function

End Class