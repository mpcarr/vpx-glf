# Getting started with GLF

Start with a working table, then add one rule at a time. GLF's VPX runtime is configured in VBScript; exported MPF YAML is useful for comparison but is not a general configuration input to VPX. See [Configuration and YAML](configuration.md).

## Choose your starting point

To build a game from an empty VPX table, follow [Build Your First GLF Game](tutorial/index.md). It includes the object checklist, a cumulative script and a complete-game checkpoint. To start from the supplied example table instead, follow [Environment Setup](environment-setup.md). Keep its working trough, plunger, switch wiring and table hooks while learning the rules system. For an existing VPX table, follow [Installation](installing.md), including the collections, timer and initialization, exit, key and option hooks.

The example table keeps game setup in `scripts/src/game/_configuration.vbs`, modes in `scripts/src/game/modes/`, and its bundled framework in `scripts/src/glf/glf.vbs`. The GLF repository's editable framework source is under `scripts/src/`; the game may bundle a different revision.

## First playable rule

1. Verify that a game starts, serves a ball and advances correctly after a drain. Use the [trough](tutorial/tutorial-trough.md) and [plunger](tutorial/tutorial-plunger.md) tutorials if needed.
2. Choose an existing switch and confirm its `<switch_name>_active` event. Use the exact name from the table.
3. Add a small mode that starts on `ball_started`, stops on `ball_will_end`, and awards points with a [variable player](variable-player.md). Call its creation sub during configuration before `Glf_Init`.
4. Rebuild the table script and restart play. Confirm one hit awards points once, including for a second player.
5. Add feedback with a [light player](light-player.md), [show](light-shows.md) or [sound](sound-system.md).

Follow the [complete mode walkthrough](guides/complete-mode.md) when you are ready to add qualification and a timed objective.

## Suggested reading order

| Learn | Guide |
| --- | --- |
| When to start, reset and stop rules | [Game and ball lifecycle](guides/game-lifecycle.md) |
| How rules communicate | [Events](events.md) |
| Who owns progress and scores | [Player state and persistence](player-state.md) |
| How components form a feature | [Complete game mode](guides/complete-mode.md) |
| How to add more balls | [Multiball walkthrough](guides/multiball.md) |
| How to present the rules | [Displays and Godot](guides/displays.md), [sound setup](sound-system.md) |
| How to investigate a failure | [Debugging](guides/debugging.md) |

Examples in these guides either identify their required table objects or use game-defined events. Rename those placeholders to your table's names before use. The guides describe the checked-in implementation; verify changes in VPX as well as reviewing the generated script.
