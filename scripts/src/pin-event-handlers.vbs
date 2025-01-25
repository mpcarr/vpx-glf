

Dim glf_lastPinEvent : glf_lastPinEvent = Null

Dim glf_dispatch_parent : glf_dispatch_parent = 0
Dim glf_dispatch_q : Set glf_dispatch_q = CreateObject("Scripting.Dictionary")

Dim glf_frame_dispatch_count : glf_frame_dispatch_count = 0
Dim glf_frame_handler_count : glf_frame_handler_count = 0

Dim glf_dispatch_queue_int : glf_dispatch_queue_int = 0

Sub DispatchPinEvent(e, kwargs)
    AddToDispatchEvents e, kwargs, 1
End Sub

Sub AddToDispatchEvents(e, kwargs, event_type)
    If glf_dispatch_await.Exists(e) Then
        glf_dispatch_await.Remove e
    End If
    glf_dispatch_await.Add e & ";" & glf_dispatch_queue_int, Array(kwargs, event_type)
    glf_dispatch_queue_int = glf_dispatch_queue_int + 1
End Sub

Function DispatchPinHandlers(e, args)
    DispatchPinHandlers = Empty
    Dim handler : handler = args(0)
    Dim event_args, retArgs
    event_args = args(1)
    glf_frame_handler_count = glf_frame_handler_count + 1
    If IsNull(event_args(0)) Then
        Set retArgs = GlfKwargs()
    Else
        On Error Resume Next
        retArgs = event_args(0)
        If Err Then 
        Set retArgs = event_args(0)
        End If
    End If
    On Error Resume Next
        glf_dispatch_current_kwargs = retArgs	
    If Err Then 
        Set glf_dispatch_current_kwargs = retArgs
    End If
    If event_args(1) = 2 Then 'Queue Event
        Set retArgs = GetRef(handler(0))(Array(handler(2), retArgs, args(2)))
        If IsObject(retArgs) Then
            If retArgs.Exists("wait_for") Then
                DispatchPinHandlers = retArgs("wait_for")
            End If
        End If
    Else
        GetRef(handler(0))(Array(handler(2), event_args(0), args(2)))
    End If
End Function

Sub RunDispatchPinEvent(eKey, kwargs)
    Dim e
    e=Split(eKey,";")(0)
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If

    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If

    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            glf_frame_dispatch_count = glf_frame_dispatch_count + 1
            If e = "timer_training_shot_add_tick" Then
                debug.print "Adding Handler for: " & e&"_"&k(1)
            End If
            glf_dispatch_handlers_await.Add e&"_"&k(1), Array(handler, kwargs, e)
        Else
            Glf_WriteDebugLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
End Sub

Sub RunAutoFireDispatchPinEvent(e, kwargs)

    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If

    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            glf_frame_dispatch_count = glf_frame_dispatch_count + 1
            'debug.print "Adding Handler for: " & e&"_"&k(1)
            'glf_dispatch_handlers_await.Add e&"_"&k(1), Array(handler, kwargs, e)
            'If SwitchHandler(handler(0), Array(handler(2), kwargs, e)) = False Then
                'debug.print e&"_"&k(1)
                GetRef(handler(0))(Array(handler(2), kwargs, e))
            'End If
        Else
            Glf_WriteDebugLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
    Glf_EventBlocks(e).RemoveAll

End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchReplayPinEvent", e
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        kwargs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
    Next
    DispatchRelayPinEvent = kwargs
    Glf_EventBlocks(e).RemoveAll
End Function

Function DispatchQueuePinEvent(e, kwargs)
    
    AddToDispatchEvents e, kwargs, 2
    Exit Function
    
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchQueuePinEvent", e & " has no listeners"
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k,i,retArgs
    Dim handlers : Set handlers = glf_pinEvents(e)
    If IsNull(kwargs) Then
        Set kwargs = GlfKwargs()
    End If
    Glf_WriteDebugLog "DispatchQueuePinEvent", e
    Dim glf_dis_events : glf_dis_events = glf_pinEventsOrder(e)
    For i=0 to UBound(glf_dis_events)
        k = glf_dis_events(i)
        Glf_WriteDebugLog "DispatchQueuePinEvent"&e, "key: " & k(1) & ", priority: " & k(0)
        'msgbox "DispatchQueuePinEvent: " & e & " , key: " & k(1) & ", priority: " & k(0)
        'msgbox handlers(k(1))(0)
        'Call the handlers.
        'The handlers might return a waitfor command.
        'If NO wait for command, continue calling handlers.
        'IF wait for command, then AddPinEventListener for the waitfor event. The callback handler needs to be ContinueDispatchQueuePinEvent.
        Set retArgs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
        If retArgs.Exists("wait_for") And i<Ubound(glf_dis_events) Then
            'pause execution of handlers at index I. 
            Glf_WriteDebugLog "DispatchQueuePinEvent"&e, k(1) & "_wait_for"
            Dim wait_for : wait_for = retArgs("wait_for")
            kwargs.Remove "wait_for" 
            AddPinEventListener wait_for, k(1) & "_wait_for", "ContinueDispatchQueuePinEvent", k(0), Array(e, kwargs, i+1)
            Exit For
            'add event listener for the wait_for event.
            'pass in the index and handlers from this.
            'in the handler for resume queue event, process from the index the remaining handlers.
        End If
    Next
    Glf_EventBlocks(e).RemoveAll
End Function


'args Array(3)
' Array(original_event, orignal_kwargs, index)
' wait_for kwargs
' event
Function ContinueDispatchQueuePinEvent(args)
    
    Dim ownProps : ownProps = args(0)
    Dim kwargs
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1)
    End If


    Dim i,key,keys,items
    keys=ownProps(0)
    items=ownProps(1)

    'Inject handlers back into dispatch
    For i=0 to UBound(keys)
        glf_dispatch_handlers_await.Add keys(i), items(i)
    Next
    Exit Function


End Function

Sub AddPinEventListener(evt, key, callbackName, priority, args)
    Dim i, inserted, tempArray
    If Not glf_pinEvents.Exists(evt) Then
        glf_pinEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_pinEvents(evt).Exists(key) Then
        glf_pinEvents(evt).Add key, Array(callbackName, priority, args)
        Sortglf_pinEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePinEventListener(evt, key)
    If glf_pinEvents.Exists(evt) Then
        If glf_pinEvents(evt).Exists(key) Then
            glf_pinEvents(evt).Remove key
            Sortglf_pinEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_pinEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_pinEventsOrder to maintain order based on priority
    If Not glf_pinEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_pinEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_pinEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_pinEventsOrder(evt)
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
        
        ' Update the glf_pinEventsOrder with the newly ordered or modified list
        glf_pinEventsOrder(evt) = tempArray
    End If
End Sub