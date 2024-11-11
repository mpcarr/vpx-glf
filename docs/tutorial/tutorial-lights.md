# Tutorial - Lights

Now that we have a basic game life cycle, we can introduce some lights management to our table. We are going to setup a light player for the GI so that it turns on when the ball starts and off when the ball ends.

The light player is used to turn lights on and off when events happen. It inherits the priority of the mode its assigned to. For more information check out the [Light Player](/vpx-gle-framework/light-player) command reference

![lights](../images/tutorial-lights.gif)


### Light Tags

Lights can be grouped together using tags. This makes it easy to change the state of many lights all at the same time. In this example we will assign all our GI (General Illumination) lights with the tag *GI*.

To do this in the vpx editor we are going to borrow the BlinkPattern property of the light object to hold our tag names. Select all of your GI lights, and under BlinkPattern type ```GI```. You can add more tags to a light using a csv format, for example ```GI,SLING_GI```

### Add Lights to GLF

To allow the GLF to control your lights you need to add them to the glf_lights collection you created during the getting started guide. In VPX open your collection manager and add all your lights to glf_lights.

### Light Player Configuration

To control your lights you need a to use a ```LightPlayer```. You also need a mode to contain your light player. Add the following configuration to your script

```
Public Sub CreateGIMode()

    With CreateGlfMode("gi_control", 1000)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_ended") 
        With .LightPlayer()
            With .Events("mode_gi_control_started")
                With .Lights("GI")
                    .Color = "ffffff"
                End With
            End With
        End With
    End With
    
End Sub

```

Lets step through the mode config. We create a new mode called gi_control, the mode will start when the ball_started event is posted by the system and will end when the ball_ended event is posted. We setup a light player within this mode that listens for the mode started event and changes the GI lights to white color.

We now need to include this mode in our game. To do that, in your ConfigureGlfDevices sub add the following

```
    Sub ConfigureGlfDevices
        Dim ball_device_plunger
        Set ball_device_plunger = (new GlfBallDevice)("plunger")

        With ball_device_plunger
            .BallSwitches = Array("sw01")
            .EjectTimeout = 2
            .MechcanicalEject = True
            .DefaultDevice = True
        End With

        CreateGIMode() '<---Add this to enable the mode
        
        'Other device config....
    End Sub
```