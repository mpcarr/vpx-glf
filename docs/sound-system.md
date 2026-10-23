# Sound setup and buses

Register sounds and buses before configuring the modes that play them. A bus groups sounds and sets default volume and simultaneous-play limits. The [sound player](sound-player.md) maps rule events to those registered sounds.

## Register a VPX sound

Import a sound into the VPX Sound Manager under the name `sfx_award`, then configure:

```vbscript
With CreateGlfSoundBus("effects")
    .SimultaneousSounds = 8
    .Volume = 0.5
End With

With CreateGlfSound("award")
    .File = "sfx_award"
    .Bus = "effects"
    .Duration = 1200
    .Loops = 0
    .EventsWhenStopped = Array("award_sound_finished")
End With
```

Replace `Duration` with the actual clip length in milliseconds. The local bus schedules bookkeeping and stopped events using that value; it does not obtain the clip length from VPX. A zero duration can therefore signal completion immediately.

Within a mode, play and stop the registered sound:

```vbscript
With .SoundPlayer()
    With .EventName("play_award")
        .Key = "award_audio"
        .Sound = "award"
    End With
    With .EventName("stop_award")
        .Key = "award_audio"
        .Sound = "award"
        .Action = "stop"
    End With
End With
```

`Sound` is the GLF registration name; `File` is the VPX sound name. Keep both the sound and playback key consistent when stopping it.

## Bus and sound reference

| Object/property | Default | Meaning |
| --- | --- | --- |
| Bus `SimultaneousSounds` | 8 | Limit for distinct locally tracked sound files |
| Bus `Volume` | 0.5 | Used when neither sound nor playback entry overrides volume |
| Bus `SystemType` | Empty | Local VPX playback; `"bcp"` delegates to the media controller |
| Sound `File`, `Bus` | Empty | Required asset and bus names |
| Sound `Volume` | Empty | Overrides the bus volume when set |
| Sound `Duration` | 0 | Local completion scheduling in milliseconds |
| Sound `Loops` | 0 | Passed to local VPX playback; use 0 for a single play |
| Sound `EventsWhenStopped` | Empty | Events emitted when the bus stops that sound |
| Sound `Priority` | 0 | Stored/exported; local playback does not use it to preempt sounds |
| Sound `MaxQueueTime` | -1 | Stored/exported; local queuing is not implemented |

Playback-entry volume overrides sound volume, which overrides bus volume. These values replace one another rather than multiplying together. Use separate music, effects and voice buses for independently configured defaults, as the example table does in `_configuration.vbs`.

The local bus tracks by sound file. Replaying the same file replaces its tracked entry and stop timer; it does not create independent instances just because playback keys differ. At capacity, the local bus's queue branch is a TODO, so do not rely on excess sounds being played later or high-priority sounds evicting lower-priority ones.

## Music and BCP audio

For repeating music, configure the sound's loop behavior and an explicit stop event when the feature ends. Test loop duration and completion behavior separately from one-shot clips. Do not use show-player loop semantics as sound-player documentation: these are different implementations.

A bus with `.SystemType = "bcp"` sends sound play/stop commands only when BCP is enabled. The media project must define the corresponding GLF sound name. The BCP path does not apply the same local duration/volume/queue logic; configure media-side playback accordingly. See [Displays and Godot](guides/displays.md).

Verify sound names, duration, repeated triggers, bus saturation, mode cleanup and the stopped event before using sound completion to advance a rule. Working examples: the example table's `scripts/src/game/modes/sounds.vbs` and `_configuration.vbs`; source: `scripts/src/devices/sound-system.vbs` and `commands/sound-player.vbs`.
