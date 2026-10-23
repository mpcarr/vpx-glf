# Displays and Godot integration

GLF supports table-light segment displays and a separate media controller connected over BCP. Use [segment displays](../segment-display.md) for VPX alphanumeric lights. Use slides and widgets for a Godot backbox presentation. Game rules remain in GLF; media assets and layouts live in the media project.

This guide follows the checked-in example table's `godot/GMCDisplay` project and GLF's BCP implementation. It is not a general installation guide for every version of GMC.

## Start from the supplied media project

1. Use the example table's `godot/GMCDisplay/project.godot`, which contains the bundled GMC add-on, `slides/`, `widgets/`, `gmc.cfg` and media resources. Keep that known combination working before upgrading the add-on.
2. On the Windows VPX machine, install the example table's `bcp-controller/VpxBcpController.msi` if the `vpx_bcp_controller.VpxBcpController` COM component is not already registered.
3. Run the media project and confirm its output reports that GMC is listening for BCP. The bundled server defaults to port 5050; its `gmc.bcp_port` configuration can override that.
4. Match GLF's `bcpPort`, then enable the table option `Glf Backbox Control Protocol`. The table must already call `Glf_Options`, `Glf_Init` and `Glf_Exit` as described in [installation](../installing.md).

## Choose how the media process starts

Set these existing globals after loading GLF and before the connection is requested; do not redeclare them with `Dim`:

```vbscript
glf_production_mode = False
bcpDebug = True
bcpPort = 5050
```

This development setup selects `ConnectToDebug` for an already running media project. To let the controller launch the editor/project instead, use `glf_production_mode = False`, `bcpDebug = False`, and supply `bcpLocalPathToGodot` and `bcpLocalPathToProject` for your machine.

In production mode, `Glf_ConnectToBCPMediaController` passes `bcpExeName` to `ConnectToBuild`. Set it to the built media executable/command expected by your controller setup. This checkout defaults it to `cGameName & "_gmc"`; the example table's bundled GLF uses a different startup implementation. Check the version actually included in your table instead of copying its defaults blindly.

The GLF monitor is a separate connection on port 5051. Enabling `Glf Monitor` does not connect the media controller on 5050.

## Play a slide

The example media project has a `base` slide. Inside your base mode:

```vbscript
With .SlidePlayer()
    With .EventName("mode_base_started")
        .Slide = "base"
        .Action = "play"
    End With
End With
```

The mode name, event name and slide resource name must match your project. The slide player sends its action, expiry, effective priority and event arguments to BCP. Mode stop sends a context clear for slides. Use an explicit slide `remove` action when removing a slide before the mode ends.

A [widget player](../widget-player.md) adds smaller overlays, such as the example table's `left_spinner` widget. GLF sends playback commands; defining a player entry does not create the Godot scene or register its resource.

## Player data and presentation lifetime

GLF sends player-add, current-player, player-variable and machine-variable messages. Bind display resources to the corresponding data in the supplied media project. Verify two-player scoring and player changes, rather than testing only player 1. After game end, live player state is cleared; final scores are retained in machine variables.

For transient effects, use a defined expiry. Do not let missing media playback block the game: a feature waiting for a presentation should have a game-side timer or skip path that can post its completion event.

## Diagnose a blank display

| Observation | Check |
| --- | --- |
| COM/controller startup error | Controller installation and the exact registered COM class |
| Media project does not connect | Listening port, running project and GLF debug/production startup branch |
| Connection works but no slide appears | Active mode, exact trigger name, resource registration and media logs |
| Correct slide but old score | Variable binding, player number and whether player-state messages arrive |
| Overlay remains after a feature | Expiry, slide context cleanup and the current widget-player limitations |

Source: `scripts/src/bcp/bcp.vbs`, connection/options code in `scripts/src/_.main.vbs`, and `commands/slide-player.vbs` / `widget-player.vbs`. Working examples: example table `scripts/src/game/modes/base.vbs`, `inlaneshot.vbs` and `godot/GMCDisplay`.
