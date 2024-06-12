Function GetPlayerState(key)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If playerState(currentPlayer).Exists(key)  Then
        GetPlayerState = playerState(currentPlayer)(key)
    Else
        GetPlayerState = Null
    End If
End Function

Function GetPlayerScore(player)
    dim p
    Select Case player
        Case 1:
            p = "PLAYER 1"
        Case 2:
            p = "PLAYER 2"
        Case 3:
            p = "PLAYER 3"
        Case 4:
            p = "PLAYER 4"
    End Select

    If playerState.Exists(p) Then
        GetPlayerScore = playerState(p)(SCORE)
    Else
        GetPlayerScore = 0
    End If
End Function

Function GetCurrentPlayerNumber()
    Select Case currentPlayer
        Case "PLAYER 1":
            GetCurrentPlayerNumber = 1
        Case "PLAYER 2":
            GetCurrentPlayerNumber = 2
        Case "PLAYER 3":
            GetCurrentPlayerNumber = 3
        Case "PLAYER 4":
            GetCurrentPlayerNumber = 4
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If IsArray(value) Then
        If Join(GetPlayerState(key)) = Join(value) Then
            Exit Function
        End If
    Else
        If GetPlayerState(key) = value Then
            Exit Function
        End If
    End If   
    Dim prevValue
    If playerState(currentPlayer).Exists(key) Then
        prevValue = playerState(currentPlayer)(key)
        playerState(currentPlayer).Remove key
    End If
    playerState(currentPlayer).Add key, value
    
    If playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue
    End If
    
    SetPlayerState = Null
End Function

Sub FirePlayerEventHandlers(evt, value, prevValue)
    If Not playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = playerEvents(evt)
    For Each k In playerEventsOrder(evt)
        GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), Array(evt,value,prevValue)))
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, callbackName, priority, args)
    If Not playerEvents.Exists(evt) Then
        playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not playerEvents(evt).Exists(key) Then
        playerEvents(evt).Add key, Array(callbackName, priority, args)
        SortPlayerEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePlayerStateEventListener(evt, key)
    If playerEvents.Exists(evt) Then
        If playerEvents(evt).Exists(key) Then
            playerEvents(evt).Remove key
            SortPlayerEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub SortPlayerEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the playerEventsOrder to maintain order based on priority
    If Not playerEventsOrder.Exists(evt) Then
        ' If the event does not exist in playerEventsOrder, just add it directly if we're adding
        If isAdding Then
            playerEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = playerEventsOrder(evt)
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
        
        ' Update the playerEventsOrder with the newly ordered or modified list
        playerEventsOrder(evt) = tempArray
    End If
End Sub

Sub EmitAllPlayerEvents()
    Dim key
    For Each key in playerState(currentPlayer).Keys()
        FirePlayerEventHandlers key, GetPlayerState(key), GetPlayerState(key)
    Next
End Sub