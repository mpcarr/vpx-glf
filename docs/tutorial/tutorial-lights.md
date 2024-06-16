# Tutorial - Lights

Now that we have a basic game life cycle, we can introduce some lights management to our table. We are going to setup a light player for the GI so that it turns on when the ball starts and off when the ball ends. And we will setup a show player to light a ball save insert.

![lights](../images/tutorial-lights.gif)

### Download the VPX file
[Tutorial - Lights VPX File](https://github.com/mpcarr/vpx-glf/raw/main/tutorial/glf_tutorial_lights.vpx)

### Light Player Configuration

For the GI add the configuration for a light player.

```
lightCtrl.AddLightTags l100, "gi"
lightCtrl.AddLightTags l101, "gi"
lightCtrl.AddLightTags l102, "gi"
lightCtrl.AddLightTags l103, "gi"

Dim light_player_base
Set light_player_base = (new LightPlayer)("base", mode_base)

light_player_base.Add "mode_base_started", Array("gi|FF0000")
```

The light player runs under our base mode and listens for the mode started event. When the mode starts, the lights tagged as **gi** will be set to Red. When the mode ends, the lights will go out.


#### Show Player Config

Next we would like our ball save to flash an insert whilst it's running. Add the configuration for the show player

```
Dim FlashColor
FlashColor = Array( _
    Array("(lights)|100|(color)"), _
    Array("(lights)|0|(color)") _
)
```

First we define a show called FlashColor. The show contains two steps, lets breakdown whats happening.

- (lights)|100|(color)
    - (lights) - A placeholder token for our lights which we will define later
    - 100 - The brightness of the light
    - (color) - A placeholder token for the color of the lights

#### Show Player

We can now define our ShowPlayer and ShowPlayerItems. This is where we define the properties of the show

```
Dim show_player_base
Set show_player_base = (new ShowPlayer)("base", mode_base)
```

#### Show Player Items

```
Dim show_ball_save_active
Set show_ball_save_active = (new ShowPlayerItem)()
With show_ball_save_active
	.Key = "ball_save"
	.Show = FlashColor
	.Loops = -1
	.Speed = 1
End With
show_ball_save_active.AddToken "lights", "l01"
show_ball_save_active.AddToken "color", "00ff00"
```

We create a show item with the key of **ball_save**, it should run the show **FlashColor**, loop until we tell it to stop and run at the default speed.
You can see we also added two tokens, one for **lights** and one for **color**. These are the tokens needed for the FlashColor show. 

```
show_player_base.Add "ball_saves_base_timer_start", show_ball_save_active
```

Above we add the show player item to the show player. It will start to play when the event **ball_saves_base_timer_start** is emitted.

We now would like our light to blink faster when the ball save is in it's hurry up state. Lets define another show player item for that

```
Dim show_ball_save_hurry
Set show_ball_save_hurry = (new ShowPlayerItem)()
With show_ball_save_hurry
	.Key = "ball_save"
	.Show = FlashColor
	.Loops = -1
	.Speed = 2
End With
show_ball_save_hurry.AddToken "lights", "l01"
show_ball_save_hurry.AddToken "color", "00ff00"
```

Similar to the last item, however this one runs at twice the default speed. We could here also change the color of the light if we wanted by adjusting the color token.

```
show_player_base.Add "ball_saves_base_hurry_up_time", show_ball_save_hurry
```

We add this item to the show player for the ball saves hurry up event.

Finally, we want to turn the light off when the ball save enters it's grace period.

```
show_player_base.Add "ball_saves_base_grace_period", "ball_save.stop"
```

We can just add a new item entry directly to the show player in the format **key.action**