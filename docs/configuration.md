# Configuration: VBScript and YAML

GLF in this repository creates devices and modes through VBScript. It can export MPF YAML from those objects. YAML light shows have a separate conversion workflow that produces VBScript shows for VPX.

| What you want to change | Where to edit | How it reaches the game |
| --- | --- | --- |
| VPX devices, rules and scoring | Table configuration and mode `.vbs` files | Calls to `CreateGlfMode`, device constructors and their properties |
| MPF configuration exported by GLF | The originating GLF configuration | Debug export writes `glf_mpf/config`, `glf_mpf/modes` and `glf_mpf/shows` |
| A light animation exported from Godot | Show `.yaml` file | Convert to `.glf.vbs`, then include it in the table script |

The current source does not provide a general YAML loader for mode and device configuration. Editing an exported mode YAML file does not change the VPX game. MPF compatibility is a project goal; a `ToYaml` method alone does not establish identical behavior in both runtimes.

## Start with a working table

Follow [installation](installing.md) for the VPX hooks and [environment setup](environment-setup.md) for the example table build workflow. In both supplied game projects, look in `scripts/src/game/_configuration.vbs` and `scripts/src/game/modes/` for game configuration. A mode creation sub must be called during configuration before `Glf_Init`; merely defining the sub does not register the mode.

## Read a mode alongside its YAML

This reduced example follows Dark Chaos's double-scoring mode. Call `CreateDoubleScoringMode` from your configuration entry point. Register `scoring_multiplier` with `Glf_SetInitialPlayerVar "scoring_multiplier", 1` during setup, and make scoring expressions use it as shown in [dynamic values](dynamic_values.md).

```vbscript
Sub CreateDoubleScoringMode
    With CreateGlfMode("double_scoring", 700)
        .StartEvents = Array("start_double_scoring")
        .StopEvents = Array("mode_base_stopping", "stop_double_scoring")

        With .Timers("double_scoring")
            .TickInterval = 1000
            .StartValue = 0
            .EndValue = 60
            With .ControlEvents()
                .EventName = "mode_double_scoring_started"
                .Action = "restart"
            End With
        End With

        With .EventPlayer()
            .Add "timer_double_scoring_complete", Array("stop_double_scoring")
        End With

        With .VariablePlayer()
            With .EventName("mode_double_scoring_started")
                With .Variable("scoring_multiplier")
                    .Action = "set"
                    .Int = 2
                End With
            End With
            With .EventName("mode_double_scoring_stopping")
                With .Variable("scoring_multiplier")
                    .Action = "set"
                    .Int = 1
                End With
            End With
        End With
    End With
End Sub
```

The corresponding YAML structure is below. This is an MPF configuration excerpt, not a file to load into VPX.

```yaml
#config_version=6
mode:
  start_events: start_double_scoring
  stop_events: mode_base_stopping, stop_double_scoring
  priority: 700

timers:
  double_scoring:
    start_value: 0
    end_value: 60
    direction: up
    tick_interval: 1000ms
    control_events:
      - event: mode_double_scoring_started
        action: restart

event_player:
  timer_double_scoring_complete:
    - stop_double_scoring

variable_player:
  mode_double_scoring_started:
    scoring_multiplier:
      action: set
      int: 2
  mode_double_scoring_stopping:
    scoring_multiplier:
      action: set
      int: 1
```

The timer and players are siblings of `mode`, not nested under it. The timer counts from 0 to 60 with a one-second interval. Completion posts `stop_double_scoring`; the stopping event restores the multiplier. `mode_base_stopping` also stops the feature if the base mode ends first.

## Export and compare

`Glf_Init` exports MPF configuration when `glf_debugEnabled` is true. The table's `Glf Debug Log` option controls that flag through `Glf_Options`. It must be enabled when initialization runs; changing the option after initialization does not itself run the exporter. Export writes files under the relative `glf_mpf` directory and can overwrite earlier exports.

Use the runtime implementation to establish behavior, then inspect its `ToYaml` method to see the exported spelling and units. For example, a timer's `.TickInterval = 1000` exports as `tick_interval: 1000ms`. Do not assume all numeric properties use the same unit: show step times are in seconds.

## Examples to explore

Paths below are relative to each game's repository. Compare a game's bundled `scripts/src/glf/glf.vbs` with the GLF version you use if behavior differs.

| Topic | Project and source | What to follow |
| --- | --- | --- |
| Scoring and multipliers | Example table: `scripts/src/game/modes/score.vbs` | Score events, dynamic expressions and per-ball resets |
| Physical multiball locks | Example table: `scripts/src/game/modes/multiball.vbs` | Lock and multiball configuration together |
| Shot progression | Example table: `scripts/src/game/modes/targetbank.vbs` | Target shots and their progression |
| Timed double scoring | Dark Chaos: `scripts/src/game/modes/double-scoring.vbs` | Mode lifecycle, timer, event player and variable player |
| Exported double scoring | Dark Chaos: `glf_mpf/modes/double_scoring/config/double_scoring.yaml` | YAML representation of the same feature |
| Scoring expressions | Dark Chaos: `scripts/src/game/modes/score.vbs` | Combining playfield, spinner and pop multipliers |
| YAML light shows | Example table: `godot/glf-example-showcreator/shows/` | Numeric time steps and light color maps |

## Source reference

In the GLF repository, `scripts/src/_.main.vbs` contains `Glf_Init`, input parsing and export orchestration. The implementations for the example above are `scripts/src/commands/mode.vbs`, `timer.vbs`, `event-player.vbs` and `variable-player.vbs`. See the individual [mode](mode.md), [timer](timer.md), [event player](event-player.md) and [variable player](variable-player.md) pages for their APIs.
