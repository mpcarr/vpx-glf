# Installation

To install and use the Game Logic Framework in your VPX table, you will need to download the lastest vpx-glf.vbs script and include it in your table script

[vpx-glf.vbs]( )

## Global Script Required Settings

The GLF requires a few global settings to function correctly. 

#### GLF Event Timer

Add a timer object to your vpx table called Glf_EventTimer. Set it to Enabled with an Interval of 100ms.

#### Global Script Settings

 - Your table name must be ```Table1```
 - ```cGameName``` must be set to your table name
 - ```BallSize``` and ```BallMass``` must be set

**Example** 

```
Const cGameName = "MyAwesomeGame"
Const BallSize = 50			'Ball diameter in VPX units; must be 50
Const BallMass = 1			'Ball mass must be 1
```

## Adding Hooks into GLF

Once the script has been added, you need to add these calls into to your table event subs.

#### Table Init

Inside ```Table1_Init``` add ```Glf_Init()```

**Example**
```
Sub Table1_Init()
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