# Tutorial 1 - Trough

The trough is special device that collects balls from the playfield when they drain and releases balls into the plunger lane. GLF supports upto an 8 ball "physical" trough. By physical we mean that the balls are never destroyed during a game. This leads to better ball tracking and physics simulation.

![trough1](../images/trough.gif)

### Download the VPX file
[Tutorial 1 - Trough VPX File](https://github.com/mpcarr/vpx-glf/raw/main/tutorial/glf_tutorial_trough.vpx)

### Setup

To setup your trough we need to create some blocker walls and trough kickers around the apron area.

![trough1](../images/tutorial-trough1.png)

The kicker switches need to named **swTrough1** - **swTrough7** plus a **Drain** kicker. **swTrough1** switch being closest to the plunger lane and **Drain** being the kicker that collects the ball from the playfield.

### Configuration

Include enough kickers to cover your games ball capacity. e.g. if your game takes 5 balls make swTrough1 - swTrough5 + Drain.