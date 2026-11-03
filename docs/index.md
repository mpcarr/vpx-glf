# Welcome to VPX Game Logic Framework

![VPX Game Logic Framework](images/glf-logo.webp){ align=right width=200 }

The VPX Game Logic Framework (GLF) is a powerful and flexible framework for creating original virtual pinball tables using the [VPX](https://github.com/vpinball/vpinball) platform. It provides a comprehensive set of game logic components that enable you to build sophisticated pinball games without writing extensive code.

## About This Framework

GLF is inspired by [The Mission Pinball Framework](https://missionpinball.org) for real and homebrew machines. Many game logic commands mirror MPF concepts. GLF configures VPX through VBScript and can export MPF YAML; compatibility depends on the component. Read [Configuration and YAML](configuration.md) before adapting examples between the two runtimes.

One of the key goals of this project is to provide a standalone set of game logic devices that doesn't require users to install and run the MPF bridge for VPX, simplifying the development process.

## Key Features

- **Virtual Devices**: Comprehensive support for flippers, bumpers, drop targets, magnets, and more
- **Game Modes**: Flexible mode system for creating complex game rules and progression
- **Event System**: Powerful event-driven architecture for game logic
- **Player Management**: Track player scores, achievements, and high scores
- **Light Control**: Advanced lighting effects and animations
- **Backbox Displays**: Options for segment displays to full LCD Displays

## Learning guides

Start with [Getting Started](getting-started.md), then follow the [game lifecycle](guides/game-lifecycle.md), [complete timed mode](guides/complete-mode.md) and [multiball walkthrough](guides/multiball.md). For presentation and recovery, see [Displays and Godot](guides/displays.md), [sound setup](sound-system.md), [ball search](ball-search.md) and [debugging](guides/debugging.md).

## Quick Start

1. **Installation**: Follow the [Installation Guide](installing.md) to set up GLF
2. **Tutorials**: Start with our [Build Your First GLF Game](tutorial/index.md) to learn the basics
3. **Configuration**: Read [Configuration and YAML](configuration.md) for a worked mode example and a guide to the two game repositories
4. **Runtime expressions**: Use [dynamic values](dynamic_values.md) for scoring multipliers and event conditions

## Getting Help

- **Community**: Reach out to Flux or Apophis on discord and ask for a link to join our community discord for support and discussion.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
