# Tutorial - Lights

Now that we have a basic game life cycle, we can introduce some lights management to our table. We are going to setup a light player for the GI so that it turns on when the ball starts and off when the ball ends.

The light player is used to turn lights on and off when events happen. It inherits the priority of the mode its assigned to. For more information check out the [Light Player](/vpx-gle-framework/light-player) command reference

![lights](../images/tutorial-lights.gif)

### Download the VPX file
[Tutorial - Lights VPX File](https://github.com/mpcarr/vpx-glf/raw/main/tutorial/glf_tutorial_lights.vpx)

### Light Player Configuration

For the GI add the configuration for a light player.

```
mode_base.LightPlayer.Add "mode_base_started", _
Array( _
	"l00|FF0000", _
	"l01|FF0000", _
	"l02|FF0000", _
	"l03|FF0000" _
	)
```

This light player will now listen for the event **mode_base_started** and then turn on lights l100, l101, l102, l103. It will also change the color to red.


### Light Tags

The goal of our light player here is to turn on the GI. We can make this easier by giving our lights tags and then using the tag in the light player.

```
lightCtrl.AddLightTags l100, "gi, leftsling"
lightCtrl.AddLightTags l101, "gi, leftsling"
lightCtrl.AddLightTags l102, "gi, rightsling"
lightCtrl.AddLightTags l103, "gi, rightsling"
```

Now we can use the tag **gi** in our light player like so.

```
mode_base.LightPlayer.Add "mode_base_started", Array("gi|FF0000")
```