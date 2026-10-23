# Widget player

The widget player sends a named widget to the BCP media controller when an event occurs in an active mode. The widget must already exist in the media project. See [Displays and Godot](guides/displays.md) for connection setup.

## Show an overlay

The example table uses a `left_spinner` widget in `godot/GMCDisplay`:

```vbscript
With .WidgetPlayer()
    With .EventName("light_right_orbit")
        .Widget = "left_spinner"
        .Priority = 10
        .Expire = 4
    End With
End With
```

Place this inside a mode and post `light_right_orbit` while it is active. `Expire` is in seconds, unlike segment display player expiry, which is in milliseconds.

| Property | Default | Behavior |
| --- | --- | --- |
| `Widget` | Empty | Name of the media resource to play |
| `Priority` | 0 | Offset added to the mode's priority |
| `Expire` | Empty | Sent as null when omitted; otherwise passed to media playback |
| `Action` | `"play"` | Stored and exported, but not used by the current runtime playback path |
| `Debug` | False | Player debug flag |

## Current runtime limits

Playback requires `useBcp = True`. The implementation calls `bcpController.PlayWidget`, which always sends action `play` and position x=0, y=0. Setting `.Action = "remove"` does not currently remove a widget through this path. There are no widget-player position or token setters in this implementation.

Deactivation removes event listeners; it does not send an explicit widget-removal command. Give temporary overlays an expiry and verify media-side context cleanup. Use the slide player's supported remove action for content that must be explicitly removed by GLF.

Source: `scripts/src/commands/widget-player.vbs` and `scripts/src/bcp/bcp.vbs`. Example: `scripts/src/game/modes/inlaneshot.vbs` in the example table; Dark Chaos also uses widgets in its base and feature modes.
