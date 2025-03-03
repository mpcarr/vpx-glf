# Installation

To install and use the Game Logic Framework in your VPX table, you will need to download the lastest vpx-glf.vbs script and include it in your table script

[vpx-glf.vbs](https://github.com/mpcarr/vpx-glf/raw/main/scripts/vpx-glf.vbs){:target="_blank"}

## Global Script Required Settings

The GLF requires a few global settings to function correctly. 

#### GLF Game Timer

Add a timer object to your vpx table called Glf_GameTimer. Set it to Enabled with an Interval of -1ms.

#### GLF Table Collection

Create two collections using the vpx collections manager (F8).

```
glf_lights
glf_switches
```

#### Global Script Settings

 - ```cGameName``` must be set to your table name
 - ```BallSize```, ```BallMass```, ```tnob```, ```lob``` and ```gBot``` must be set
 - ```tablewidth``` and ```tableheight``` must be set.

**Example** 

```
Const cGameName = "MyAwesomeGame"
Const BallSize = 50			'Ball diameter in VPX units; must be 50
Const BallMass = 1			'Ball mass must be 1
Const tnob = 6              'Total playable balls (Balls in Trough)
Const lob = 2               'Total non playable balls (Captive Balls)
Dim gBot                    'Collection used for ball tracking
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

```

## Configure Devices

You'll need an entry point to configure your devices and enable modes. Create a sub called ```ConfigureGlfDevices```

```
Sub ConfigureGlfDevices
    'Device confguration goes here.
End Sub
```

## Adding Hooks into GLF

You need to add these calls into to your table event subs.

#### Table Init

Inside ```Table1_Init``` add your ```ConfigureGlfDevices``` sub call and ```Glf_Init()```

**Example**
```
Sub Table1_Init()
    ConfigureGlfDevices()
	Glf_Init()
End Sub
```

#### Table Exit

Inside ```Table1_Exit``` add ```Glf_Exit()```

**Example**
```
Sub Table1_Exit()
	Glf_Exit()
End Sub
```

#### Table Keys

Inside your ```Table1_KeyDown``` and ```Sub Table1_KeyUp``` Subs, add the ```Glf_KeyDown``` and ```Glf_KeyUp``` calls.

**Example**

```
Sub Table1_KeyDown(ByVal keycode)
    Glf_KeyDown(keycode)
End Sub

Sub Table1_KeyUp(ByVal keycode)
    Glf_KeyUp(keycode)
End Sub
```

#### Table Options

Inside ```Table1_Options``` add ```Glf_Options(eventId)```

**Example**

```
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True

    Glf_Options(eventId)

    If eventId = 3 Then DisableStaticPreRendering = False
End Sub

```


## Next Steps


[Setting up your trough](../tutorial/tutorial-trough/)