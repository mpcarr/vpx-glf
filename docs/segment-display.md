# Segment display devices

A segment display maps characters to VPX light objects. Configure the device globally before `Glf_Init`, then use the [segment display player](segment-display-player.md) inside modes to supply text.

## Create an eight-character display

Prepare the table's lights first. A `14Segment` display uses 15 lights per character, including punctuation; an eight-character display needs 120 lights named `p1_seg1` through `p1_seg120`. Register those lights with GLF through the normal light collection setup. Their ordering must match the segment mapping; the example table supplies a working layout.

```vbscript
With CreateGlfLightSegmentDisplay("player1")
    .SegmentType = "14Segment"
    .SegmentSize = 8
    .LightGroup = "p1_seg"
    .UpdateMethod = "stack"
    .DefaultColor = "ffffff"
    .UseDotsForCommas = True
End With
```

Set segment type and size before the light-group property because these setters calculate the light mapping. The prefix is a naming convention for numbered light objects, not a VPX collection name.

## Settings

| Property | Default | Purpose |
| --- | --- | --- |
| `SegmentType` | Empty | `"14Segment"` or `"7Segment"`; set explicitly |
| `SegmentSize` | 0 | Number of character positions |
| `LightGroup` | Empty | Prefix for consecutive numbered lights |
| `LightGroups` | Empty array | Multiple prefixes for combined displays; use the example table's layout as a reference |
| `UpdateMethod` | `"replace"` | `"stack"` allows priority-based text entries and restoration of underlying text |
| `DefaultColor` | `"ffffff"` | Default display color |
| `IntegratedCommas`, `IntegratedDots` | False | Punctuation handling within character positions |
| `UseDotsForCommas` | False | Use dots to represent commas |
| `DefaultTransitionUpdateHz` | 30 | Transition update rate |
| `ExternalFlexDmdSegmentIndex` | -1 | Optional external FlexDMD segment offset |
| `ExternalB2SSegmentIndex` | -1 | Optional B2S segment offset |

A `7Segment` display uses eight lights per position. External output indices refer to the external display layout, not the VPX light suffixes. In the example table, player 1 uses FlexDMD index 0 and B2S index 1; reproduce the complete external layout setup before applying those values to another table.

## Check the mapping

Display a fixed string first, then a score. Verify each character position and punctuation. If text is scrambled, check segment order and light names before changing event logic. If temporary text never reveals the score again, confirm `UpdateMethod = "stack"`, the entry keys, priorities and expiry.

Source: `scripts/src/devices/light-segment_display.vbs`. Working example: the alphanumeric-display section of `vpx-example-glf/scripts/src/game/_configuration.vbs`, including combined `all` and `p1p2` displays.
