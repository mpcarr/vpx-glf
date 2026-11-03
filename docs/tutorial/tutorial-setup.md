# 1. Prepare the empty table

**Goal:** prepare a new table and a single script that subsequent chapters extend.

Use VPX with the scripting features used by this GLF checkout, including `Table1.Option`. Save a new table as `GLFFirstGame.vpx` and name the table object `Table1`. Begin with simple, visible geometry: perimeter walls, two flippers, an open drain below them, and a shooter lane on the right that feeds the upper playfield. Keep the lane wide enough for a 50-unit-diameter ball, with clearance rather than a tight fit.

If starting from VPX's default new-table template, remove its sample game script and old ball creation/drain logic. Keep only the physical pieces you want and rename them to the [object list](index.md#one-naming-scheme). You do not need artwork, ROM code, sound, B2S, Godot or a script watcher to complete this tutorial.

## Editor setup

Create the four collections `glf_lights`, `glf_switches`, `glf_slingshots` and `glf_spinners` in the collection manager. Add the timers and backglass TextBox from the object list. Put `ScoreText` on the backglass and test in desktop view so it is visible. Leave the display timer disabled until chapter 5 if its callback is not yet in your script.

Use the framework build `scripts/vpx-glf.vbs` from your GLF checkout and put it in VPX's `Scripts` folder, alongside the installed standard scripts. Load it only once. Define the table constants before loading GLF because GLF reads `tnob` and `cGameName` during script execution.

## Initial script

Paste this into the table's script editor. Leave one empty configuration sub for chapter 3:

```vbscript
Const cGameName = "GLFFirstGame"
Const BallSize = 50
Const BallMass = 1
Const tnob = 1
Const lob = 0
Dim gBot
Dim tablewidth : tablewidth = Table1.Width
Dim tableheight : tableheight = Table1.Height

ExecuteGlobal GetTextFile("controller.vbs")
ExecuteGlobal GetTextFile("core.vbs")
ExecuteGlobal GetTextFile("vpx-glf.vbs")

Sub Table1_Init()
    ConfigureGlfDevices
    Glf_Init Table1
End Sub

Sub Table1_Exit()
    Glf_Exit
End Sub

Sub Table1_KeyDown(ByVal keycode)
    Glf_KeyDown keycode
    If keycode = PlungerKey Then Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal keycode)
    Glf_KeyUp keycode
    If keycode = PlungerKey Then Plunger.Fire
End Sub

Sub Table1_OptionEvent(ByVal eventId)
    Glf_Options eventId
End Sub

Sub ConfigureGlfDevices()
    ' Chapter 3 adds the plunger; later chapters append more calls.
End Sub
```

`controller.vbs` and `core.vbs` come with the VPX installation. Do not suppress load errors: a missing GLF or core script must be fixed before continuing. GLF provides `Glf_GameTimer_Timer`, trough handlers and generated switch handlers; do not paste copies of them into this script.

`tnob = 1` means one physical ball. It does **not** mean a one-ball game: the same ball is reused over three player turns. `lob = 0` means no captive balls. `Glf_Init` creates the physical ball and fills `gBot`; do not create another ball in your own initialization.

`Glf_KeyDown`/`Glf_KeyUp` handle GLF inputs, while the explicit `Plunger.Pullback`/`Plunger.Fire` calls move the VPX plunger. Defining an eject callback alone does not wire keyboard plunger operation.

## Checkpoint

Before Play, finish the trough and plunger in chapters 2–3. At that point the table must have every referenced object, a default GLF ball device and one call to `Glf_Init`. Keep `Glf Monitor` and `Glf Backbox Control Protocol` off for this tutorial. If the table reports a missing object or unsupported `Option` method, resolve the setup/version mismatch before adding rules.

[Previous](index.md) · [Tutorial map](index.md) · [Next](tutorial-trough.md)
