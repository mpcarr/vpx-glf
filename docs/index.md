# Welcome to VPX Game Logic Framework

This project is a framework for making original virtual pinball tables using the [VPX](https://github.com/vpinball/vpinball) platform.
It is based on [The Mission Pinball Framework](https://missionpinball.org) for real and homebrew machines. Many of the game logic commands mirror the MPF commands as one of the goals of this project is to have an interchangeable config between vpx and mpf. Another goal of this project is to have a standalone set of game logic devices that doesn't not require the end user to install and run the mpf bridge for vpx.

## Installation

Check out the [Installing GLF Guide](installing.md) to get started.

## Features

### Game Logic Devices

A well designed Virtual Pinball table contains many of the real world devices such as scoops, vuks (vertical up kickers), plungers, diverters, drop targets e.t.c. These devices are machine wide devices that need to respond to the ball interacting with them. The classes found in the framework mimic these device behaviours by monitoring and dispatching events which other parts of the system can respond to.

- [Ball Device](ball-device.md)
- [Diverters](diverter.md)
- [Drop Targets](drop-target.md)

### Game Logic Commands

To program your game, the framework provides a set of classes to manage the game life-cycle, player state, player score, modes, timers, light shows and manage your DMD or LCD game display.

- [Ball Save](ball-save.md)
- [Counter](counter.md)
- [Event Player](event-player.md)
- [Light Player](light-player.md)
- [Mode](mode.md)
- [Multball Locks](multiball-locks.md)
- [Multiball](multiball.md)
- [Show Player](show-player.md)
- [Timer](timer.md)

### Events System

### Player State