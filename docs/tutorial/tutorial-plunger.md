# Tutorial - Plunger

The plunger is a ball device which holds a ball until either the plunger is pulled back and released or a button is pressed to launch the ball. 
It is the default ball device that is used to get balls onto the playfield.

### Setup

To setup your plunger, you'll need a switch that is positioned so that the ball is resting on it whilst in the plunger lane.

![trough1](../images/plunger1.png)

In the example we have named the switch **sw01**.

### Configuration

To configure your plunger device add the code below to your Configure Glf Devices Sub.

```
Sub ConfigureGlfDevices
    
    ' Plunger
    With CreateGlfBallDevice("plunger")
        .BallSwitches = Array("sw01")
        .EjectTimeout = 2000
        .MechanicalEject = True
        .DefaultDevice = True
		.EjectCallback = "PlungerEjectCallback"
    End With

End Sub
```

The above configuration sets up the plunger as a mechcanical plunger, meaning the ball won't automatically fire out of the plunger lane, you will have to pull and release the plunger to fire it. We have also set the ball device as the default device (this is the ball device which recieves a ball from the trough). The eject timeout setting tells the device how long to wait until deeming the ball successfully exited the device, e.g. If the ball didn't reach the end of the plunger lane and fell back within 2 seconds, the ball would not have exited successfully.

## Next Steps

[Setting up your lights](../tutorial-lights/)