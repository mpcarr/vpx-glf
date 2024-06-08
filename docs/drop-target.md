# Drop Target

Drop targets in vpx are implemented via Rothbauerw's DropTarget code. This device extends Roth's Drop Target class to include Knockdown and Reset Events

# Example

```
Dim drop1 : Set drop1 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False, Array("ball_started"," machine_reset_phase_3"))
```