# Tutorial - Shows

Shows are a list of light state changes to produce simple or complex light shows. The most simple light show would be a blinking light. We will setup a show to blink a light and then use a show player to play the show when an event happens in our game.

![lights](../images/tutorial-lights.gif)

### Download the VPX file
[Tutorial - Shows VPX File](https://github.com/mpcarr/vpx-glf/raw/main/tutorial/glf_tutorial_lights.vpx)

### Shows

Below is a show called **FlashColor**. It's a reuseable show in that we can use tokens to swap out the values. 

```
Dim FlashColor
FlashColor = Array( _
    Array("(lights)|100|(color)"), _
    Array("(lights)|0|(color)") _
)
```

The show contains two steps, lets breakdown whats happening.

- (lights)|100|(color)
    - (lights) - A placeholder token for our lights which we will define later
    - 100 - The brightness of the light
    - (color) - A placeholder token for the color of the lights

You can create shows with hardcoded light names and colors like below. You can also mix and match with tokens.

```
Dim FlashL01
FlashL01 = Array( _
    Array("l01|100|FF0000"), _
    Array("l01|0|FF0000") _
)
```

### Show Player

Each mode has a show player which is used to play the shows we've defined.

```
Glf_ShowTokens_Add "ball_save_tokens", "lights", "l01"
Glf_ShowTokens_Add "ball_save_tokens", "color", "00ff00"

							'Event						  	 'Key 		   'Show       'Loops		'Speed		'Tokens
mode_base.ShowPlayer().Add 	"ball_saves_base_timer_start", 	 "ball_save", 	FlashColor, -1, 		1,			"ball_save_tokens"
```

It will start to play when the event **ball_saves_base_timer_start** is emitted. We create a show item with the key of **ball_save**, it should run the show **FlashColor**, loop until we tell it to stop and run at the default speed.

You can see we also added two tokens, one for **lights** and one for **color**. These are the tokens needed for the FlashColor show. 

We now would like our light to blink faster when the ball save is in it's hurry up state. Lets define another show player item for that

```
mode_base.ShowPlayer().Add 	"ball_saves_base_hurry_up_time", "ball_save", 	FlashColor, -1, 		2,			"ball_save_tokens"
```

We add this item to the show player for the ball saves hurry up event. Similar to the last item, however this one runs at twice the default speed.

Finally, we want to turn the light off when the ball save enters it's grace period.

```
mode_base.ShowPlayer().StopShow "ball_saves_base_grace_period",  "ball_save"
```

### Full Config Example

```
'Shows
Dim FlashColor : FlashColor = Array(Array("(lights)|100|(color)"), Array("(lights)|0|(color)"))

Glf_ShowTokens_Add "ball_save_tokens", "lights", "l01"
Glf_ShowTokens_Add "ball_save_tokens", "color", "00ff00"

							'Event						  	 'Key 		   'Show       'Loops		'Speed		'Tokens
mode_base.ShowPlayer().Add 	"ball_saves_base_timer_start", 	 "ball_save", 	FlashColor, -1, 		1,			"ball_save_tokens"
mode_base.ShowPlayer().Add 	"ball_saves_base_hurry_up_time", "ball_save", 	FlashColor, -1, 		2,			"ball_save_tokens"
mode_base.ShowPlayer().StopShow "ball_saves_base_grace_period",  "ball_save"
```