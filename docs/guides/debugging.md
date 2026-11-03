# Debugging game rules

Follow one failing interaction from physical input to its final effect. Check the running table's bundled GLF version first: both example games carry their own `scripts/src/glf/glf.vbs`, which may differ from the framework source you are reading.

## Enable useful logging

Enable the table option `Glf Debug Log`, and use `Glf Debug Log Level` when more detail is needed. Make sure the table calls `Glf_Options`. Logs are written under the relative `glf_logs` folder with a filename beginning with `cGameName` and a timestamp.

Enable `.Debug = True` on the mode and the component being investigated where those objects expose it. Global logging alone does not make every component emit its optional detail. Debug initialization also exports MPF configuration; see [Configuration and YAML](../configuration.md) before editing exported files.

`Glf Monitor` uses a separate BCP connection on port 5051 and requires the monitor/controller setup. Ordinary file logging is enough to begin investigating a rule.

## Trace one hit

1. Find the emitted switch event, such as `s_RampHit_active`. If absent, check the VPX object name, collection or specialized-device registration, generated handlers and tilt state.
2. Find `mode_<name>_started`. Confirm the mode creation sub ran before `Glf_Init`, its start event happened and it has not already stopped.
3. Find the listener for the rule event. Check exact spelling and whether conditions reference initialized variables or existing devices.
4. Check priorities and intermediate events. An earlier handler may change a value, disable a device or stop a mode before a later action runs.
5. Follow the effect to its target: a registered show, sound, light, ball device or media resource. A valid event with an unknown asset still cannot produce the expected result.

## Common failures

| Symptom | Likely checks |
| --- | --- |
| New code has no effect | Rebuild the actual table script; confirm the table loads it and restart play |
| YAML edit has no effect in VPX | Mode YAML is exported; change the VBScript source. Show YAML must be converted again |
| Mode starts but initialization is missed | Use `mode_<name>_started` for that mode's players instead of the event that starts the mode |
| Score awards twice | Two active scoring handlers, duplicate switch dispatch, or overlapping unconditional/conditional mappings |
| Expression fails at configuration time | Quote runtime expressions, such as `.Int = "1000 * current_player.multiplier"` |
| Counter never counts | Give the counter an explicit enable event after its mode activates |
| Progress belongs to the wrong player | Separate player credit from shared machine/device state; verify index conventions |
| Next ball never serves | Find the queued event and missing `WaitFor` completion; also check trough/plunger state |
| Saved initial value seems ignored | The INI reload may override it; inspect the named variable's saved value |
| Sound completion fires too soon | Check the registered sound's duration and loop settings |
| Media is blank | Check BCP connection, mode context and resource names in the media project |

The example table's watcher concatenates `.vbs` changes; it does not convert show YAML. Its production concatenation also includes `cached-functions.vbs`. Use the development workflow while changing expressions, and check the cache/build process if the generated script does not reflect the intended change.

## Verify cleanup and player boundaries

After fixing the normal path, exercise timeout, mode cancellation, tilt, drain, next player and a new game. Look for timers still posting events, repeating shows or sounds, held balls, stale shot states and duplicate listeners. A successful first hit is not enough to establish a feature's lifecycle.

For a reproducible report, record the table/GLF revision, active modes, event sequence, player/ball, relevant variable values, expected behavior and the smallest configuration excerpt that reproduces the issue.

Source: `scripts/src/logs.vbs`, `pin-event-handlers.vbs`, `_.main.vbs` and the component under investigation. Related: [events](../events.md), [game lifecycle](game-lifecycle.md), [player state](../player-state.md).
