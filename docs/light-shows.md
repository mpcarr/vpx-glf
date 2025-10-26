# Light Shows Tutorial

This tutorial will walk you through setting up a workflow for designing **pinball light shows in Godot** and exporting them for use in **Visual Pinball X (VPX)**.

By using the [GMC plugin](https://github.com/missionpinball/mpf-gmc) and [GMC Toolbox](https://github.com/missionpinball/gmc-toolkit), you’ll be able to visually lay out lights, animate shows in the Godot Editor, and export them as YAML files that can be used in your VPX table using.

## Step 1: Set Up Your Godot Project

Follow the excellent guides on installing the mpf-gmc and gmc-toolkit from the above repositories.

## Step 6: Position Your Lights

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

## Step 7: Animate Your Show

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

## Step 8: Export Your Show

1. Go back to the *GMC Toolbox* panel
2. Click **Refresh Animations**
3. Select your animation from the dropdown list
4. Set the FPS and strip options
5. Click **Generate Show**

> 📝 The show will be saved to your project's `/shows/` folder as a `.yaml` file.

This file can now be used as a light show in your VPX project using MPF!

---

## Step 9: Use the Show in VPX

In your MPF config (e.g., mode or show YAML), reference the show file you created:

```yaml
- play_show:
    show: intro_flash
