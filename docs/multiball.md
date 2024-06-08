# Multiball

## Example

```
Dim waterfall_mb
Set waterfall_mb = (new Multiball)("waterfall", "multiball_locks_waterfull" ,mode_waterfall_mb)

With waterfall_mb
    .EnableEvents = Array("mode_waterfall_mb_started")
    .DisableEvents = Array("mode_waterfall_mb_ended")
    .StartEvents = Array("multiball_locks_waterfall_full")
End With
```

See multiball locks for how to count locked multiballs. In this example we create a waterfall multiball, the multiball will start when the multiball locks emits that it is full.