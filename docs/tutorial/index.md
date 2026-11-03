# Build your first GLF game

Build a small VPX table with one physical ball, three turns per player, a manual plunger, two flippers, one scoring rollover and general illumination. By the end of the core tutorial you can start a game, play, score, drain three times and start another game. A simple backglass text box displays the current player, ball and score; no external media controller is required.

This is a cumulative build. Keep one table script and one `ConfigureGlfDevices` sub. Each chapter tells you exactly which objects and subs to add. The first three chapters form the initial runnable setup; finish them before pressing Play.

## Core path

| Chapter | What you build | Checkpoint |
| --- | --- | --- |
| [1. Empty table and script](tutorial-setup.md) | Required objects, collections and GLF loading | Names and hooks are ready |
| [2. Physical trough](tutorial-trough.md) | Ball storage and return path | One ball can return to the release kicker |
| [3. Plunger and first serve](tutorial-plunger.md) | Shooter lane, tracking and launch | Start serves a ball; drain serves the next |
| [4. Flippers](tutorial-flippers.md) | Both flipper callbacks | Flippers work during a ball and release at ball end |
| [5. Score and display](tutorial-scoring.md) | One rollover, score rule and text display | One pass scores 1,000 for the current player |
| [6. Lights and complete-game test](tutorial-lights.md) | GI controlled by a mode | Complete three turns and start a new game |

The [completed core script](../assets/tutorial/minimal-game.vbs) is a comparison/reference file for chapters 1–6. It does not create VPX geometry or replace the object setup below. It has been checked against GLF source, but the new tutorial has not yet been play-tested in VPX. Existing `.vpx` files in the repository's `tutorial/` folder are older geometry references, not validated snapshots of this new script.

## One naming scheme

| VPX object | Type / settings | Collection |
| --- | --- | --- |
| `Table1` | Table with a sloped playfield, outlanes and shooter lane | — |
| `swTrough1` | Capturing kicker at the trough release position | None; GLF owns its handlers |
| `Drain` | Capturing kicker reached by all drained balls | None; GLF owns its handlers |
| `Plunger` | Mechanical plunger facing up the shooter lane | None |
| `s_plunger` | Trigger covered by a resting ball at the plunger | `glf_switches` |
| `LeftFlipper`, `RightFlipper` | Flippers with mirrored start/end angles | None |
| `s_score` | Rollover trigger on the playfield | `glf_switches` |
| `l_gi` | Light, initially off, BlinkPattern `GI` | `glf_lights` |
| `Glf_GameTimer` | Timer, Enabled, Interval -1 | None; GLF owns its timer sub |
| `TutorialDisplayTimer` | Timer, Enabled, Interval 100 | None |
| `ScoreText` | Backglass TextBox, visible in desktop view | None |

Create empty `glf_slingshots` and `glf_spinners` collections as well. Keep collection event generation disabled: GLF generates the individual switch handlers. Do not add the trough or drain kickers to `glf_switches` and do not write duplicate `_Hit` or `_UnHit` subs for registered objects.

## Extend the working game

After chapter 6, add these in order or select the independent optional devices:

1. [Event routing](tutorial-event-player.md): separate a physical switch from its award.
2. [Ball save](tutorial-ballsave.md): save one early drain and automatically launch the replacement.
3. [Shot progress](tutorial-shots.md): light an insert and score a completed objective.
4. [Timed opportunity](tutorial-timer.md): light a timed award window.
5. [Optional scoop](tutorial-ball-devices.md): capture and return a ball.
6. [Optional autofire device](tutorial-autofire.md): add a simple kickback.

Advanced modes, multiball, sound, Godot, tilt and high scores follow after this core loop. See [the learning guides](../getting-started.md). This tutorial uses local VBScript configuration; exported YAML is explained separately in [Configuration and YAML](../configuration.md).

[Tutorial map](index.md) · [Next](tutorial-setup.md)
