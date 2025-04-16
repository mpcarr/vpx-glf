# VPX Game Logic Framework

A powerful framework for creating original virtual pinball tables using the VPX platform. This framework provides a comprehensive set of game logic components that make it easy to build complex pinball machines without extensive coding.

## Documentation

For detailed documentation, please visit our [official documentation site](https://mpcarr.github.io/vpx-glf/).

## Overview

The VPX Game Logic Framework is based on The Mission Pinball Framework (MPF) for real and homebrew machines. Many of the game logic commands mirror the MPF commands, with one of the project's goals being to have an interchangeable configuration between VPX and MPF. Another goal is to provide a standalone set of game logic devices that doesn't require end users to install and run the MPF bridge for VPX.

## Key Features

- **Virtual Devices**: Ball devices, flippers, auto-fire devices, diverters, drop targets, and magnets
- **Game Modes**: Comprehensive mode system for managing game states
- **Event System**: Robust event handling for game interactions
- **Player Management**: Track player states, scores, and achievements
- **Light Control**: Advanced lighting effects and animations
- **Sound Management**: Coordinate sound effects with game events
- **High Score Tracking**: Persistent high score storage and display
- **Achievement System**: Track and reward player accomplishments
- **Shot Management**: Configure shots, shot groups, and shot profiles
- **Multiball Handling**: Manage multiball modes and locks
- **Timer System**: Precise timing control for game events
- **Tilt Management**: Configure tilt behavior and recovery

## Getting Started

Check out the [Installation Guide](https://mpcarr.github.io/vpx-glf/installing/) to get started with the framework.

## Tutorials

The documentation includes step-by-step tutorials for implementing common pinball features:

- [Trough](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-trough/)
- [Plunger](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-plunger/)
- [Flippers](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-flippers/)
- [Auto Fire Devices](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-autofire/)
- [Lights](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-lights/)
- [Ball Devices](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-ball-devices/)
- [Ball Save](https://mpcarr.github.io/vpx-glf/tutorial/tutorial-ballsave/)

## Game Logic Components

The framework provides a wide range of game logic components:

### Virtual Devices
- Ball Device
- Flipper
- Auto Fire Device
- Diverter
- Drop Target
- Magnet

### Modes
- Ball Save
- Combo Switches
- DOF Player
- Event Player
- Extra Ball
- High Score
- Light Player
- Multiball
- Multiball Locks
- Queue Event Player
- Random Event Player
- Shot
- Shot Group
- Shot Profile
- Sequence Shot
- Show Player
- Slide Player
- Sound Player
- State Machine
- Timer
- Timed Switches
- Tilt
- Variable Player

## License

MIT License

Copyright (c) 2023 VPX Game Logic Framework

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- Based on concepts from The Mission Pinball Framework
- Thanks to the VPX community for their support and feedback