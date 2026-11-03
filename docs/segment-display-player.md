# Segment display player

A segment display player adds, removes or changes text on a registered [segment display device](segment-display.md). It runs within a mode and adds the mode's priority to each text entry's priority.

## Show a temporary award

After creating a display named `player1`, put this inside a mode:

```vbscript
With .SegmentDisplayPlayer()
    With .EventName("award_jackpot")
        With .Display("player1")
            .Text = "JACKPOT"
            .Key = "jackpot_message"
            .Priority = 10
            .Expire = 2000
            .Flashing = "all"
        End With
    End With
    With .EventName("clear_jackpot")
        With .Display("player1")
            .Key = "jackpot_message"
            .Action = "remove"
        End With
    End With
End With
```

Use `UpdateMethod = "stack"` on the display to restore lower-priority text after the message expires or is removed. Keys are scoped by mode and display. A matching key replaces or removes the same entry even when the triggering event differs.

## Text and dynamic values

Text uses brace placeholders rather than a bare numeric expression:

```vbscript
.Text = "{current_player.score}"
```

Literal text such as `"JACKPOT"` does not need additional embedded quotes. The setter converts formatted text into a GLF input expression. Test live updates and player changes with the intended reference; do not assume an arbitrary expression automatically establishes subscriptions to every value it contains.

## Settings and actions

| Property | Default | Purpose |
| --- | --- | --- |
| `Text` | Null | Literal text or brace placeholders |
| `Action` | `"add"` | Add text or perform one of the actions below |
| `Key` | Empty | Distinguishes entries within the mode/display |
| `Priority` | 0 | Added to mode priority |
| `Expire` | 0 | Milliseconds; zero means no expiry timer |
| `Flashing` | `"off"` | Display flashing mode, including `"all"`, `"match"` and `"mask"` |
| `FlashMask` | Empty | Mask for selective flashing |
| `Color` | `RGB(255,255,255)` | Text color supplied to the display |

Actions implemented by the callback include `add`, `remove`, `flash`, `flash_match`, `flash_mask`, `off` and `set_color`. Here `off` turns flashing off; use `remove` to remove the text entry.

`Transition()` and `TransitionOut()` expose `TransitionType`, `Direction` and `Text`. The transition defaults are `push` and `right`; push and cover implementations are in `devices/segment-transitions/`. Start with static text before adding transitions.

When the mode deactivates, the player removes its text entries and updates affected stacks. Verify the resulting display when several modes target the same device, particularly when they use different priorities or overlapping combined displays.

Source: `scripts/src/commands/segment-display-player.vbs`. Examples: the example table's base and attract modes, and Dark Chaos's `scripts/src/game/modes/double-scoring.vbs`.
