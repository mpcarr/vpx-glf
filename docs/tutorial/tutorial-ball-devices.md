# Tutorial - Ball Devices

Ball devices are anything in your game that capture balls from the playfield. These can be scoops, vuks (vertical up kickers), troughs, plungers, multiball locks e.t.c. 

[See Ball Device Config Reference](../../ball-device)

Below is an image of Cyber Race pinball. Apart from the trough and plunger it has four circled ball devices. There are three vuks and one scoop. All of these devices collect, hold balls and release them.

![cyberrace-balldevices](../images/cyberrace-balldevices.png)

### Setup

A ball device will need a switch to dermine when a ball has entered and some way of holding the ball physically. There are many ways to do this in VPX such as using a kicker object or building an enclosure with walls. 

### Configuration

To configure your device add the config below to your Configure Glf Devices Sub.

```
Sub ConfigureGlfDevices
    Dim ball_device_scoop
    Set ball_device_scoop = (new GlfBallDevice)("scoop")

    With ball_device_scoop
        .BallSwitches = Array("sw04")
        .EjectTimeout = 2
        .EjectStrength = 100
        .EjectAngle = 0
        .EjectPitch = 90
    End With

    'Other device config....
End Sub
```

The above configuration sets up the scoop device with a switch called **sw04**. When the ball enters the scoop and the game is expecting the ball to enter, the ball will be held until released by the system. Once released it will be ejected by kicking the ball stright up. The scoop should have a physical wall above it to direct the ball back into play.

### Example Direction Configuration

| **Direction**         | **Angle (degrees)** | **Pitch (degrees)** | **Notes**                         |
|------------------------|---------------------|----------------------|------------------------------------|
| Forward               | 0                   | 0                    | Straight ahead, flat.             |
| Right                 | 90                  | 0                    | Flat, to the right.               |
| Backward              | 180                 | 0                    | Straight back, flat.              |
| Left                  | 270                 | 0                    | Flat, to the left.                |
| Upward                | Any                 | 90                   | Vertical or diagonal upward.      |
| Downward              | Any                 | -90                  | Vertical or diagonal downward.    |
| Diagonal Up-Right     | 45                  | Positive (30)        | Combines rightward and upward.    |
| Diagonal Down-Left    | 225                 | Negative (-30)       | Combines leftward and downward.   |
