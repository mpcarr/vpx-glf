# Build a multiball feature

A multiball lock tracks qualification; a multiball device manages balls in play and shoot-again. A physical ball device knows where a ball actually is. Keep those responsibilities separate, especially in a multi-player game.

## Start with virtual locks

Virtual locks credit progress without retaining a physical ball. This setup needs a working trough and auto-launch plunger. Call `CreateDemoMultiball` during configuration. Post `qualify_demo_locks` from a completed shot group or other qualification rule, then post `demo_lock_hit` from the shot that collects each lock.

```vbscript
Sub CreateDemoMultiball
    With CreateGlfMode("demo_mb", 1000)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")

        With .MultiballLocks("demo_locks")
            .EnableEvents = Array("qualify_demo_locks")
            .DisableEvents = Array("disable_demo_locks")
            .LockEvents = Array("demo_lock_hit")
            .ResetEvents = Array("prepare_demo_locks")
            .BallsToLock = 2
        End With

        With .Multiballs("demo")
            .StartEvents = Array("start_demo_mb")
            .BallCount = 3
            .BallCountType = "total"
            .ShootAgain = 15000
            .HurryUp = 3000
            .GracePeriod = 2000
        End With

        With .EventPlayer()
            .Add "qualify_demo_locks", Array("prepare_demo_locks")
            .Add "multiball_lock_demo_locks_full", Array("start_demo_mb")
            .Add "multiball_demo_started", Array("disable_demo_locks", "enable_demo_jackpots")
            .Add "multiball_demo_ended", Array("disable_demo_jackpots", "restart_demo_qualification")
        End With
    End With
End Sub
```

The lock's reset listener exists while the lock is enabled. Posting a separate `prepare_demo_locks` event after qualification allows it to reset the current player's lock count after enabling. Use plain event names in `LockEvents` and `ResetEvents`: this implementation registers those arrays directly rather than parsing event conditions or priority suffixes.

Three balls with `BallCountType = "total"` means a target of three in play, not three additional balls. Do not manually increment `glf_BIP` to launch multiball.

## Add jackpots and requalification

`enable_demo_jackpots`, `disable_demo_jackpots` and `restart_demo_qualification` above are hooks for your rules, not built-in GLF commands. One approach is a dedicated jackpot mode that starts on `multiball_demo_started` and stops on `multiball_demo_ended`, `ball_will_end` or `tilt`. Its variable player can score a jackpot switch only while that mode is active.

Keep the qualification rules disabled during multiball. On `restart_demo_qualification`, reset the qualification shots and allow the player to earn `qualify_demo_locks` again. For a feature that should be available only once per ball, leave qualification disabled until the next ball instead.

## Replace virtual locks with physical locks

Configure two [ball devices](../ball-device.md), for example `kicker1` and `kicker2`, with switches, eject callbacks and capacities appropriate to your table. In the lock configuration, replace `LockEvents` with:

```vbscript
.LockDevices = Array("kicker1", "kicker2")
```

In the multiball configuration, add:

```vbscript
.BallLocks = Array("kicker1", "kicker2")
```

The lock claims entering balls and schedules replacement balls when appropriate; the multiball releases the listed physical devices when it starts. Disable competing ball holds for those devices while locking. Qualification still determines when the lock is enabled.

Player lock credit is stored as `<lock_name>_locked_balls`. Physical occupancy belongs to the device and can differ from a player's credit. The example table's `scripts/src/game/modes/multiball.vbs` demonstrates stealable locks using machine occupancy variables, shot states and a state machine. Study that complete implementation before adding shared physical locks between players.

`LockedBallCountingStrategy` has a setter and export support, but the current runtime lock calculation does not branch on it. Do not assume selecting an MPF strategy name changes GLF behavior.

## End and shoot-again behavior

The multiball emits `multiball_<name>_started`, `_shoot_again`, `_hurry_up`, `_grace_period`, `_shoot_again_ended`, `_ball_lost` and `_ended` during its lifecycle. The drain-count handler posts `_ended` when the remaining global ball count would be one (or the game is no longer running). That normally returns to single-ball play; it does not end the player's ball.

Timing values above are milliseconds. Test drains during shoot-again, around its expiry, and after it ends. Test two players with different lock credit and physically occupied locks, and verify a new game does not inherit stale qualification. Check [Multiball](../multiball.md), [Multiball Locks](../multiball-locks.md) and source `scripts/src/commands/multiball.vbs` and `multiball-locks.vbs` for additional settings.
