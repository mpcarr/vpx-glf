# Counter

A simple counter to monitor things like the number of times a target for switch has been hit.

## Example

```
Dim counter_pop_hits
Set counter_pop_hits = (new Counter)("pop_hits", mode_super_pops)

With counter_pop_hits
    .EnableEvents = Array("mode_super_pops_started")
    .CountEvents = Array("sw30_active", "sw31_active", "sw32_active")
    .CountCompleteValue = 20
    .DisableOnComplete = True
    .ResetOnComplete = True
    .EventsWhenComplete = Arrry("super_pops_qualified")
    .PersistState = True
End With
```

In the above example we created a counter called **counter_pop_hits** that belongs to the mode called **mode_super_pops**. The counter is enabled when the mode starts and will increase its count value whenever one of the three pop bumbers are hit (sw30,sw31,sw32). Once complete, the counter will disable and reset its count and emit the event *super_pops_qualified*. The Persist State property means the counter will not reset between balls.

## Required Setings

### Name
```String```

The name of this device. Events emitted from the device will be in the format **name_counter**

### Mode
```String```

This is the mode the counter belongs to. 