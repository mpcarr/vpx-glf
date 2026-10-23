# Light Shows Tutorial

This tutorial will walk you through setting up a workflow for designing **pinball light shows in Godot** and exporting them for use in **Visual Pinball X (VPX)**.

By using the [GMC plugin](https://github.com/missionpinball/mpf-gmc) and [GMC Toolbox](https://github.com/missionpinball/gmc-toolkit), you’ll be able to visually lay out lights, animate shows in the Godot Editor, and export them as YAML files, then convert those files into GLF shows for your VPX table.

## Step 1: Set Up Your Godot Project

Follow the excellent guides on installing the mpf-gmc and gmc-toolkit from the above repositories.

## Step 2: Position Your Lights

1. Open `playfield.tscn`
2. In the scene tree, you'll see a list of `GMCLight` nodes
3. Use the **Move Mode** (crosshair icon at top) to drag and position each light over your playfield image
4. The warning icons will disappear once each light is positioned
5. Optionally, you can customize each light's:
   - Shape
   - Size (in inches)
   - Rotation

6. Save the scene

---

## Step 3: Animate Your Show

1. In the scene tree, add an `AnimationPlayer` node as a child of `GMCPlayfield`
2. Select the root `GMCPlayfield` node
   - In the Inspector, under `GMCPlayfield > Animation Player`, assign your `AnimationPlayer`

3. Use shapes (`ColorRect`, `TextureRect`, gradients, etc.) to design your animation
4. Add keyframes for:
   - `Modulate` (color)
   - `Position`, `Rotation`, `Scale` (movement)
5. Name your animation something meaningful (e.g., `intro_flash`, `sweep_left`)

💡 Tip: Use the Animation panel at the bottom of Godot to set durations, create tracks, and preview your animation in real time.

---

## Step 4: Export Your Show

1. Go back to the *GMC Toolbox* panel
2. Click **Refresh Animations**
3. Select your animation from the dropdown list
4. Set the FPS and strip options
5. Click **Generate Show**

> 📝 The show will be saved to your project's `/shows/` folder as a `.yaml` file.

Convert this file to VBScript before including it in your VPX table, as described below.

---

## Step 5: Convert the YAML for GLF

The following commands belong to the **example table repository**, not the GLF framework repository. Its `scripts/mpf-show-convert.js` converter reads YAML steps and writes a `.glf.vbs` file beside each input file.

1. Copy your exported YAML into `scripts/src/glf/yamlshows/` in the example table.
2. From the example table's `scripts` directory, install dependencies if needed with `npm install`.
3. Run the conversion:

```sh
npm run convert-shows
```

The converter expects a list of steps with numeric, absolute `time` values in seconds and a `lights` map of light names to hex color strings. Quote colors so YAML preserves leading zeroes and does not interpret them as numbers.

```yaml
- time: 0
  lights:
    l_example: "ff0000"
- time: 0.5
  lights:
    l_example: "000000"
- time: 1
  lights:
    l_example: "000000"
```

Save this as `intro_flash.yaml` and replace `l_example` with a light from your table. Conversion creates `intro_flash.glf.vbs` containing `CreateGlfShow("intro_flash")`. The converter handles `time` and `lights`; it does not implement the full MPF show format, such as duration-only steps or other show players.

For a grayscale animation whose color should be supplied at playback, use `npm run convert-shows-tokens` instead. That converter derives intensity from the red channel and substitutes `(color)` for the output color. Supply a `color` token when playing the resulting show. Both commands write the same output filenames, so choose one conversion style for each input folder.

## Step 6: Build and play the show

The example table's Grunt configuration includes `src/glf/yamlshows/*.vbs` in the assembled table script. Run `npm run concat-vpx`, or use the script watcher described in [environment setup](environment-setup.md). Rerun the converter whenever you change the YAML; the watcher does not perform YAML conversion.

Inside an existing mode, add a show player entry:

```vbscript
With .ShowPlayer()
    With .EventName("play_intro_flash")
        .Key = "key_intro_flash"
        .Show = "intro_flash"
        .Loops = 0
        .Speed = 1
    End With
End With
```

Post `play_intro_flash` while that mode is active. The show name must match the YAML filename without its extension. For a tokenized show, add the following inside the `EventName` block:

```vbscript
With .Tokens()
    .Add "color", "ff0000"
End With
```

See [Show Player](show-player.md) for playback settings and [Configuration and YAML](configuration.md) for the distinction between show conversion and mode configuration export.
