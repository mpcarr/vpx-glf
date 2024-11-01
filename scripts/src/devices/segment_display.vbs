Class SegmentDisplay

    Public default Function init(name)

        glf_segment_displays.Add name, Me
        Set Init = Me
    End Function

    Public Sub AddTextEntry(text, color, flashing, flash_mask, transition, transition_out, priority, key)

    End Sub

    Public Sub RemoveTextByKey(key)

    End Sub

    Public Sub SetFlashing(flash_type)

    End Sub

    Public Sub SetFlashingMask(mask)

    End Sub

    Public Sub SetColor(color)

    End Sub

End Class