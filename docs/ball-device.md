# `GlfBallDevice` Class Documentation

The `GlfBallDevice` class is used to manage and configure ball devices in a pinball machine. The class provides methods and properties to handle ball movement, events, and device-specific configurations.

[See Ball Devices](../tutorial/tutorial-ball-devices)

---


## Properties

### **BallSwitches**
```Array<String>```  
An array of switch names that detect ball presence in the device.

---

### **DefaultDevice**
```Boolean```  
Specifies whether this device is the default ball device. If set to `True`, this device becomes the global plunger reference.

---

### **EjectAllEvents**
```Array<String>```  
An array of event names that will trigger the ejection of all balls in the device.

---

### **EjectCallback**
```String```  
The name of the callback function to execute when a ball is ejected.

---

### **EjectTargets**
```Array<String>```  
An array of target event names for ejection.

---

### **EjectTimeout**
```Integer```  
The timeout for an ejection attempt, in seconds.

---

### **MechcanicalEject**
```Boolean```  
Specifies whether the device uses mechanical ejection.

---

### **PlayerControlledEjectEvent**
```Array<String>```  
event name triggered by player-controlled eject attempts.

---

## Example Usage

Here’s an example of how to configure a ball device:

```
With CreateGlfBallDevice("scoop")
    .BallSwitches = Array("sw04")
    .EjectTimeout = 2000
    .EjectCallback = "ScoopEjectCallback"
End With
```