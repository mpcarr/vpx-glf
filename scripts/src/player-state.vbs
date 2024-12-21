Function GetPlayerState(key)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If glf_playerState(glf_currentPlayer).Exists(key)  Then
        GetPlayerState = glf_playerState(glf_currentPlayer)(key)
    Else
        GetPlayerState = False
    End If
End Function

Function GetPlayerStateForPlayer(player, key)
    dim p
    Select Case player
        Case 0:
            p = "PLAYER 1"
        Case 1:
            p = "PLAYER 2"
        Case 2:
            p = "PLAYER 3"
        Case 3:
            p = "PLAYER 4"
    End Select

    If glf_playerState.Exists(p) Then
        GetPlayerStateForPlayer = glf_playerState(p)(key)
    Else
        GetPlayerStateForPlayer = False
    End If
End Function

Function Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            Getglf_currentPlayerNumber = 0
        Case "PLAYER 2":
            Getglf_currentPlayerNumber = 1
        Case "PLAYER 3":
            Getglf_currentPlayerNumber = 2
        Case "PLAYER 4":
            Getglf_currentPlayerNumber = 3
    End Select
End Function

Function Getglf_PlayerName(player)
    Select Case player
        Case 0:
            Getglf_PlayerName = "PLAYER 1"
        Case 1:
            Getglf_PlayerName = "PLAYER 2"
        Case 2:
            Getglf_PlayerName = "PLAYER 3"
        Case 3:
            Getglf_PlayerName = "PLAYER 4"
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If IsArray(value) Then
        If Join(GetPlayerState(key)) = Join(value) Then
            Exit Function
        End If
    Else
        If VarType(GetPlayerState(key)) <> vbBoolean Then
            If GetPlayerState(key) = value Then
                Exit Function
            End If
        End If
    End If   
    Dim prevValue
    If glf_playerState(glf_currentPlayer).Exists(key) Then
        prevValue = glf_playerState(glf_currentPlayer)(key)
        glf_playerState(glf_currentPlayer).Remove key
    End If
    glf_playerState(glf_currentPlayer).Add key, value
    If glf_debug_level = "Debug" Then
        Dim p,v : p = prevValue : v = value
        If IsNull(prevValue) Then
            p=""
        End If
        If IsNull(value) Then
            v=""
        End If
        Glf_WriteDebugLog "Player State", "Variable "& key &" changed from " & CStr(p) & " to " & CStr(v)
    End If
    If Not IsNull(glf_debugBcpController) Then
        glf_monitor_player_state = glf_monitor_player_state & "{""key"": """&key&""", ""value"": """&value&"""},"
    End If
    If glf_playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue, -1
    End If
    
    SetPlayerState = Null
End Function

Function SetPlayerStateByPlayer(key, value, player)

    If IsArray(value) Then
        If Join(GetPlayerStateForPlayer(player, key)) = Join(value) Then
            Exit Function
        End If
    Else
        If GetPlayerStateForPlayer(player, key) = value Then
            Exit Function
        End If
    End If   
    Dim prevValue, player_name
    player_name = Getglf_PlayerName(player)
    If glf_playerState(player_name).Exists(key) Then
        prevValue = glf_playerState(player_name)(key)
        glf_playerState(player_name).Remove key
    End If
    glf_playerState(player_name).Add key, value
    
    If glf_playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue, player
    End If
    
    SetPlayerStateByPlayer = Null
End Function

Sub FirePlayerEventHandlers(evt, value, prevValue, player)
    If Not glf_playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = glf_playerEvents(evt)
    For Each k In glf_playerEventsOrder(evt)
        If handlers(k(1))(3) = player or handlers(k(1))(3) = Getglf_currentPlayerNumber() Then
            GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), Array(evt,value,prevValue)))
        End If
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, player, callbackName, priority, args)
    If Not glf_playerEvents.Exists(evt) Then
        glf_playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_playerEvents(evt).Exists(key) Then
        glf_playerEvents(evt).Add key, Array(callbackName, priority, args, player)
        Sortglf_playerEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePlayerStateEventListener(evt, key)
    If glf_playerEvents.Exists(evt) Then
        If glf_playerEvents(evt).Exists(key) Then
            glf_playerEvents(evt).Remove key
            Sortglf_playerEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_playerEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_playerEventsOrder to maintain order based on priority
    If Not glf_playerEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_playerEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_playerEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_playerEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the glf_playerEventsOrder with the newly ordered or modified list
        glf_playerEventsOrder(evt) = tempArray
    End If
End Sub

Sub EmitAllglf_playerEvents()
    Dim key
    For Each key in glf_playerState(glf_currentPlayer).Keys()
        FirePlayerEventHandlers key, GetPlayerState(key), GetPlayerState(key), -1
    Next
End Sub