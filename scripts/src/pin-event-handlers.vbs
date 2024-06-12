

Dim glf_lastPinEvent : glf_lastPinEvent = Null

Sub DispatchPinEvent(e, kwargs)
    If Not pinEvents.Exists(e) Then
        'debugLog.WriteToLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = pinEvents(e)
    'debugLog.WriteToLog "DispatchPinEvent", e
    For Each k In pinEventsOrder(e)
        'debugLog.WriteToLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs))
    Next
End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not pinEvents.Exists(e) Then
        'debugLog.WriteToLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = pinEvents(e)
    'debugLog.WriteToLog "DispatchReplayPinEvent", e
    For Each k In pinEventsOrder(e)
        'debugLog.WriteToLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        kwargs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs))
    Next
    DispatchRelayPinEvent = kwargs
End Function

Sub AddPinEventListener(evt, key, callbackName, priority, args)
    Dim i, inserted, tempArray
    If Not pinEvents.Exists(evt) Then
        pinEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not pinEvents(evt).Exists(key) Then
        pinEvents(evt).Add key, Array(callbackName, priority, args)
        SortPinEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePinEventListener(evt, key)
    If pinEvents.Exists(evt) Then
        If pinEvents(evt).Exists(key) Then
            pinEvents(evt).Remove key
            SortPinEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub SortPinEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the pinEventsOrder to maintain order based on priority
    If Not pinEventsOrder.Exists(evt) Then
        ' If the event does not exist in pinEventsOrder, just add it directly if we're adding
        If isAdding Then
            pinEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = pinEventsOrder(evt)
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
        
        ' Update the pinEventsOrder with the newly ordered or modified list
        pinEventsOrder(evt) = tempArray
    End If
End Sub