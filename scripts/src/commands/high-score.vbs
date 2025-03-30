Function EnableGlfHighScores()
    Dim high_score_mode : Set high_score_mode = CreateGlfMode("glf_high_scores", 80)
    high_score_mode.StartEvents = Array("game_ending")
    high_score_mode.StopEvents = Array("high_score_complete")
    high_score_mode.UseWaitQueue = True
    Dim high_score : Set high_score = (new GlfHighScore)(high_score_mode)
    high_score.Debug = True
    high_score_mode.HighScore = high_score
    Set glf_highscore = high_score
	Set EnableGlfHighScores = glf_highscore
End Function

Class GlfHighScore

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    private m_categories
    private m_defaults
    private m_vars
    private m_award_slide_display_time
    private m_enter_initials_timeout
    private m_reset_high_scores_events
    private m_highscores
    Private m_initials_needed
    Private m_current_initials

    Public Property Get Name() : Name = "high_score" : End Property

    Public Property Get Categories()
        Set Categories = m_categories
    End Property

    Public Property Get Defaults(name)
        Dim new_default : Set new_default = CreateObject("Scripting.Dictionary")
        m_defaults.Add name, new_default
        Set Defaults = m_defaults(name)
    End Property

    Public Property Get Vars(name)
        m_vars.Add name, CreateObject("Scripting.Dictionary")
        Set Vars = m_vars(name)
    End Property

    Public Property Let AwardSlideDisplayTime(input)
        Set m_award_slide_display_time = CreateGlfInput(input)
    End Property

    Public Property Let EnterInitialsTimeout(input)
        Set m_enter_initials_timeout = CreateGlfInput(input)
    End Property

    Public Property Let ResetHighScoreEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_high_scores_events.Add newEvent.Raw, newEvent
        Next
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "high_score_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_award_slide_display_time = CreateGlfInput(4000)
        Set m_enter_initials_timeout = CreateGlfInput(20000)
        Set m_categories = CreateObject("Scripting.Dictionary")
        Set m_defaults = CreateObject("Scripting.Dictionary")
        Set m_vars = CreateObject("Scripting.Dictionary")
        Set m_highscores = CreateObject("Scripting.Dictionary")
        Set m_initials_needed = CreateObject("Scripting.Dictionary")
        Set m_reset_high_scores_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "high_score", Me)
        ReadHighScores()
        Set Init = Me
	End Function

    Public Sub Activate()
        
        'Run High Scores.
        'Check if any player variable beats the stored value for that category.
        'Ask for intiital if we dont have them.
        'display award
        'write scores.
        'finish
        Log "Activating"
        Dim category, p
        
        For Each category in m_categories.Keys()
            'eg score.
            ' get the value for each player
            Dim player_values()
            Dim player_numbers()
            ReDiM player_values(UBound(glf_playerState.Keys()))
            ReDiM player_numbers(UBound(glf_playerState.Keys()))
            For p=0 to UBound(glf_playerState.Keys())
                player_values(p) = GetPlayerStateForPlayer(p, category)
                player_numbers(p) = p+1
            Next
            'Sort the values high to low.
            Dim i, j, temp
            For i = 0 To UBound(player_values) - 1
                For j = i + 1 To UBound(player_values)
                    If player_values(j) > player_values(i) Then
                        temp = player_values(i)
                        player_values(i) = player_values(j)
                        player_values(j) = temp

                        temp = player_numbers(i)
                        player_numbers(i) = player_numbers(j)
                        player_numbers(j) = temp
                    End If
                Next
            Next

            'msgbox player_values(0)
            
            For i = 0 To UBound(player_values)
                Dim position
                'msgbox "UBound(m_categories(category)): " & UBound(m_categories(category))
                For position=0 to UBound(m_categories(category))
                    Dim high_score
                    'msgbox "Category: " & category
                    high_score = m_highscores(category)(CStr(position+1))("value")
                    high_score = CLng(high_score)
                    ''msgbox "HighScore: " & ">" & high_score & "<"
                    ''msgbox "PlayerScore: " & CInt(player_values(i))
                    'msgbox typename(high_score)
                    'msgbox typename(player_values(i))
                    If player_values(i) > high_score Then
                        'msgbox "setting new high score"
                        Log "Setting new high score"

                        'Shift everything down, knocking the bottom score off.
                        Dim shift_index, hs_item, hs_tmp
                        Set hs_tmp = GlfKwargs()
                        With hs_tmp
                            .Add "label", ""
                            .Add "player_name", ""
                            .Add "value", 0
                        End With
                        Set hs_item = GlfKwargs()
                        With hs_item
                            .Add "label", ""
                            .Add "player_name", ""
                            .Add "value", 0
                        End With
                        For shift_index=position+1 to UBound(m_categories(category))
                            If shift_index>position+1 Then
                                hs_item("label") = hs_tmp("label")
                                hs_item("player_name") = hs_tmp("player_name")
                                hs_item("value") = hs_tmp("value")
                            Else
                                hs_item("label") = m_highscores(category)(CStr(shift_index))("label")
                                hs_item("player_name") = m_highscores(category)(CStr(shift_index))("player_name")
                                hs_item("value") = m_highscores(category)(CStr(shift_index))("value")
                            End If
                            hs_tmp("label") = m_highscores(category)(CStr(shift_index+1))("label")
                            hs_tmp("player_name") = m_highscores(category)(CStr(shift_index+1))("player_name")
                            hs_tmp("value") = m_highscores(category)(CStr(shift_index+1))("value")

                            ''msgbox "hs_item" & hs_item("value")
                            'msgbox "hs_tmp" & hs_tmp("value")

                            m_highscores(category)(CStr(shift_index+1))("label") = hs_item("label")
                            m_highscores(category)(CStr(shift_index+1))("player_name") = hs_item("player_name")
                            m_highscores(category)(CStr(shift_index+1))("value") = hs_item("value")
                        Next

                        'new score
                        m_highscores(category)(CStr(position+1))("value") = player_values(i)
                        m_highscores(category)(CStr(position+1))("player_name") = ""
                        m_highscores(category)(CStr(position+1))("player_num") = player_numbers(i)
                        m_highscores(category)(CStr(position+1))("award") = m_categories(category)(position)
                        m_highscores(category)(CStr(position+1))("category") = category
                        m_highscores(category)(CStr(position+1))("position") = position + 1
                        If Not m_initials_needed.Exists(i+1) Then
                            m_initials_needed.Add i+1, m_highscores(category)(CStr(position+1))
                        End If
                        Exit For
                    Else
                        'msgbox "Player Score " & player_values(i) & " is not greater than the high score: " & high_score
                    End If
                Next
            Next
        Next

        'Ask for Initials
        If Ubound(m_initials_needed.Keys())>-1 Then
            'msgbox "Asking For Initials"
            'msgbox m_initials_needed.Items()(0)("position")
            Log "Asking for Initials"
            m_current_initials = 0
            AddPinEventListener "text_input_high_score_complete", "text_input_high_score_complete", "HighScoreEventHandler", m_priority, Array("initials_complete", Me, m_initials_needed.Items()(0))
            DispatchPinEvent "high_score_enter_initials", m_initials_needed.Items()(0)
            SetDelay "enter_initials_timeout", "HighScoreEventHandler", Array(Array("initials_complete", Me, m_initials_needed.Items()(0)), Null), m_enter_initials_timeout.Value
        Else
            'No New High Scores, End Mode
            Log "No High Score, Ending"
            'msgbox "No High Score, Ending"
            DispatchPinEvent "high_score_complete", Null
        End If

    End Sub

    Public Sub Deactivate()
        
    End Sub

    Public Sub InitialsInputComplete(initials_item, kwargs)

        'Show Award.
        Log "Initials Complete"
        'msgbox "Initials Complete"
        RemoveDelay "enter_initials_timeout"
        RemovePinEventListener "text_input_high_score_complete", "text_input_high_score_complete"

        'm_highscores(initials_item("category"))(CStr(initials_item("position")))("player_name") = text

        Dim text : text = ""
        If Not IsNull(kwargs) Then
            If kwargs.Exists("text") Then
                text = kwargs("text")
            End If
        End If

        Dim keys, key
        keys = m_highscores.Keys()
        For Each key in keys
            Dim s
            For Each s in m_highscores(key).Keys()
                If m_highscores(key)(s)("player_num") = initials_item("player_num") Then
                    'msgbox "Setting Player " & m_current_initials+1 & " Name to >" & text & "<"
                    m_highscores(key)(s)("player_name") = text
                End If
            Next
        Next

        Log "Show Award"
        DispatchPinEvent "high_score_award_display", initials_item
        DispatchPinEvent initials_item("award") & "_award_display", initials_item
        DispatchPinEvent initials_item("category") & "_award_display", initials_item
    
        SetDelay "award_display_time", "HighScoreEventHandler", Array(Array("award_display_complete", Me), Null), m_award_slide_display_time.Value

    End Sub

    Public Sub AwardDisplayComplete()

        Log "Award Complete"
        If Ubound(m_initials_needed.Keys())>m_current_initials Then
            Log "Asking for Next Initials"
            m_current_initials = m_current_initials + 1
            AddPinEventListener "text_input_high_score_complete", "text_input_high_score_complete", "HighScoreEventHandler", m_priority, Array("initials_complete", Me, m_initials_needed.Items()(m_current_initials))
            DispatchPinEvent "high_score_enter_initials", m_initials_needed.Items()(m_current_initials)
            SetDelay "enter_initials_timeout", "HighScoreEventHandler", Array(Array("initials_complete", Me, m_initials_needed.Items()(m_current_initials)), Null), m_enter_initials_timeout.Value
        Else
            Log "Writing High Scores"
            Dim keys, key
            keys = m_highscores.Keys()
            Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")
            For Each key in keys
                Dim s, i
                i=1
                'msgbox "Key Count" & ubound(m_highscores(key).Keys())
                For Each s in m_highscores(key).Keys()
                    'msgbox s
                    tmp.Add key & "_" & i & "_label", m_categories(key)(i-1)
                    tmp.Add key & "_" & i &"_name", m_highscores(key)(s)("player_name")
                    tmp.Add key & "_" & i &"_value", m_highscores(key)(s)("value")
                    i=i+1
                Next
                WriteHighScores "HighScores", tmp
            Next
            Log "Ending"
            DispatchPinEvent "high_score_complete", Null
        End If        
        'End Mode
    End Sub

    Public Sub WriteDefaults()
        ReadHighScores()
        dim key, keys, i
        keys = m_defaults.Keys()
        Dim tmp : Set tmp = CreateObject("Scripting.Dictionary")
        For Each key in keys
            For i=0 to UBound(m_defaults(key).Keys())
                If m_highscores.Exists(key) Then
                    If Not m_highscores(key).Exists(CStr(i+1)) Then
                        tmp.Add key & "_" & i+1 &"_label", m_categories(key)(i)
                        tmp.Add key & "_" & i+1 &"_name", m_defaults(key).Keys()(i)
                        tmp.Add key & "_" & i+1 &"_value", m_defaults(key)(m_defaults(key).Keys()(i))
                    End If
                Else
                    tmp.Add key & "_" & i+1 & "_label", m_categories(key)(i)
                    tmp.Add key & "_" & i+1 &"_name", m_defaults(key).Keys()(i)
                    tmp.Add key & "_" & i+1 &"_value", m_defaults(key)(m_defaults(key).Keys()(i))
                End If
            Next
            WriteHighScores "HighScores", tmp
        Next  
        ReadHighScores()
    End Sub

    Private Sub WriteHighScores(section, scores)
        Dim objFSO, objFile, arrLines, line, inSection, foundSection
        Dim outputLines, key
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        If objFSO.FileExists(CGameName & "_glf.ini") Then
            Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
            arrLines = Split(objFile.ReadAll, vbCrLf)
            objFile.Close
        Else
            arrLines = Array()
        End If
        
        outputLines = ""
        inSection = False
        foundSection = False
        
        For Each line In arrLines
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase(section))
                foundSection = foundSection Or inSection
            End If
            
            If inSection And InStr(line, "=") > 0 Then
                key = Trim(Split(line, "=")(0))
                If scores.Exists(key) Then
                    line = key & "=" & scores(key)
                    scores.Remove key
                End If
            End If
    
            If line = "" And inSection Then
                ' Add remaining keys in the section
                For Each key In scores.Keys
                    outputLines = outputLines & key & "=" & scores(key) & vbCrLf
                Next
                scores.RemoveAll
            End If
            If line <> "" Then
                outputLines = outputLines & line & vbCrLf
            End If
        Next
        
        If Not foundSection Then
            outputLines = outputLines & "["&section&"]" & vbCrLf
            For Each key In scores.Keys
                outputLines = outputLines & key & "=" & scores(key) & vbCrLf
            Next
        End If
        
        Set objFile = objFSO.CreateTextFile(CGameName & "_glf.ini", True)
        objFile.Write outputLines
        objFile.Close
        Glf_ReadMachineVars("HighScores")
    End Sub

    Sub ReadHighScores()
        Dim objFSO, objFile, arrLines, line, inSection
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        If Not objFSO.FileExists(CGameName & "_glf.ini") Then Exit Sub
        
        Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
        arrLines = Split(objFile.ReadAll, vbCrLf)
        objFile.Close
        
        m_highscores.RemoveAll()

        inSection = False
        Dim current_name
        For Each line In arrLines
            line = Trim(line)
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase("HighScores"))
            ElseIf inSection And InStr(line, "=") > 0 Then
                Dim parts : parts = Split(line, "=")
                Dim key_parts : key_parts = Split(parts(0), "_")

                Dim category : category = Trim(key_parts(0))
                Dim position : position = Trim(key_parts(1))
                Dim attr : attr = Trim(key_parts(2))
                Dim current_label
                If m_categories.Exists(category) Then
                    If Not m_highscores.Exists(category) Then
                        m_highscores.Add category, CreateObject("Scripting.Dictionary")
                    End If
                    Select Case attr
                        Case "label":
                            current_label = parts(1)
                        Case "name":
                            current_name = parts(1)
                        Case "value":
                            Dim kwargs : Set kwargs = GlfKwargs()
                            With kwargs
                                .Add "label", current_label
                                .Add "player_name", current_name
                                .Add "value", parts(1)
                            End With
                            'msgbox category & "," & position & ", " & current_label & ", " & current_name & ", " & parts(1)
                            m_highscores(category).Add CStr(position), kwargs
                    End Select
                End If
            End If
        Next
    End Sub
    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_high_score", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml : yaml = ""
        ToYaml = yaml
    End Function

End Class

Function HighScoreEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim high_score : Set high_score = ownProps(1)
    Select Case evt
        Case "activate"
            high_score.Activate
        Case "deactivate"
            high_score.Deactivate
        Case "initials_complete"
            high_score.InitialsInputComplete ownProps(2), kwargs
        Case "award_display_complete"
            high_score.AwardDisplayComplete
    End Select
    If IsObject(args(1)) Then
        Set HighScoreEventHandler = kwargs
    Else
        HighScoreEventHandler = kwargs
    End If
End Function

Class GlfHighScoreItem
	Private m_slide, m_action, m_expire, m_max_queue_time, m_method, m_priority, m_target, m_tokens
    
    Public Property Get Slide(): Slide = m_slide: End Property
    Public Property Let Slide(input)
        m_slide = input
    End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        m_action = input
    End Property

    Public Property Get Expire(): Expire = m_expire: End Property
    Public Property Let Expire(input)
        m_expire = input
    End Property

    Public Property Get MaxQueueTime(): MaxQueueTime = m_max_queue_time: End Property
    Public Property Let MaxQueueTime(input)
        m_max_queue_time = input
    End Property

    Public Property Get Method(): Method = m_method: End Property
    Public Property Let Method(input)
        m_method = input
    End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input)
        m_priority = input
    End Property

    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input)
        m_target = input
    End Property

    Public Property Get Tokens(): Tokens = m_tokens: End Property
    Public Property Let Tokens(input)
        m_tokens = input
    End Property

	Public default Function init()
        m_action = "play"
        m_slide = Empty
        m_expire = Empty
        m_max_queue_time = Empty
        m_method = Empty
        m_priority = Empty
        m_target = Empty
        m_tokens = Empty
        Set Init = Me
	End Function

End Class
