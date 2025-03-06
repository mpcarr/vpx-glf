# `GlfFlipper` Class Documentation

The `GlfFlipper` class is used to manage and configure flippers in a pinball machine. The class provides configuration for enable / disabling flippers.

[See Flippers tutorial](../tutorial/tutorial-flippers)

---

## Properties

### **Switche**
```String```  
The name of the switch used to control the flipper

---

### **EnableEvents**
```Array<GlfEvent>```  
A list of events that when dispatched will enable the flipper

---

### **DisableEvents**
```Array<GlfEvent>```  
A list of events that when dispatched will disable the flipper

---

### **ActionCallback**
```String```  
The name of the function that will be called with the flipper activate or deactivates

---

### **Debug**
```Boolean```  
Debugs the flipper

---

## Example Usage

```
With CreateGlfFlipper("left")
    .Switch = "s_left_flipper"
    .ActionCallback = "LeftFlipperAction"
    .DisableEvents = Array("ball_will_end")
    .EnableEvents = Array("ball_started")
End With
```