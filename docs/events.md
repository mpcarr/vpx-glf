# Events System

The VPX GLF works on an events based system. All the devices emit and listen to events posted via the commands below

## Add Pin Event Listener

Used to add an event listener. Parameters are:

- Event: The event you want to listen to
- Key: A unqiue key for this event
- Callback: The function to call when this event fires
- Priority: The priority the callback should be fired in relation to other callbacks for this event. (Higher priority will be called first)
- Arguments: Anything to be passed along to the callback function when this event fires.

## Example

```
AddPinEventListener "sw01_active", "switch01Active", "HandleSwitchHit", 1000, Null

Sub HandleSwitchHit(args)
    Dim listenerArgs, dispatchArgs
    listenerArgs = args(0)
    dispatchArgs = args(1)
End Sub
```
When the event callback is called, a param will be passed along. This is an array where the first item is the args set with AddPinEventListener and the second item is the args sent with the DispatchPinEvent call.


## Remove Pin Event Listener

Used to remove a pin event listener

```
RemovePinEventListener "sw01_active", "switch01Active"
```

## Dispatch Pin Event

Used to dispatch a pin event to the system

```
DispatchPinEvent "sw01_active", Null
```

## Dispatch Relay Pin Event

Used to dispatch a special pin event to the system where the argument passed with the pin event is relayed and returned to each function.

```
DispatchRelayPinEvent "sw01_active", 1
```

Listeners to these relay events must return the argument so that the next function can received the updated value

```
Function HandleSwitchHit(args)
    Dim listenerArgs, dispatchArgs
    listenerArgs = args(0)
    dispatchArgs = args(1)
    dispatchArgs = 2
    HandleSwitchHit = dispatchArgs
End Function
```

In the above example, the initial event dispatched passed 1 as the argument, the HandleSwitchHit received this, changed it to a 2 and returned that value. The next event callback to fire would then recieved the value 2 instead of the initial 1.