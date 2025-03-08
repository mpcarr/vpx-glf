# Tutorial - Auto Fire Devices

[See Auto Fire Devices Config Reference](../../autofire)

## Introduction


## **Configuration**


```
Sub ConfigureGlfDevices
  
    ' Slingshots
    With CreateGlfAutoFireDevice("left_sling")
        .Switch = "s_LeftSlingshot"
        .ActionCallback = "LeftSlingshotAction"
        .DisabledCallback = "LeftSlingshotDisabled"
        .EnabledCallback = "LeftSlingshotEnabled"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started","enable_flippers")
    End With

    With CreateGlfAutoFireDevice("right_sling")
        .Switch = "s_RightSlingshot"
        .ActionCallback = "RightSlingshotAction"
        .DisabledCallback = "RightSlingshotDisabled"
        .EnabledCallback = "RightSlingshotEnabled"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started","enable_flippers")
    End With

End Sub
```