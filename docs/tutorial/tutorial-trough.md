# Tutorial 1 - Trough

The trough is special device that collects balls from the playfield when they drain and releases balls into the plunger lane. GLF supports upto an 8 ball "physical" trough. By physical we mean that the balls are never destroyed during a game. This leads to better ball tracking and physics simulation.

### Download the VPX file
[Tutorial 1 - Trough VPX File](https://github.com/mpcarr/vpx-glf/raw/main/tutorial/glf_tutorial1_trough.vpx)

### Setup

To setup your trough we need to create some blocker walls and trough kickers around the apron area.

![trough1](../images/tutorial-trough1.png)

The kicker switches need to named **swTrough1** - **swTrough8** with the **swTrough1** switch being closest to the plunger lane and **swTrough8** being the ball drain kicker.

### Configuration

The default settings for the trough are a 6 ball setup. To change this there will be a table option automatically added by GLF. Press F12 to open the VPX options menu and navigate to the table options. You will see a setting for trough capacity. You can change this here whilst a game is not in progress.

![trough2](../images/tutorial-trough2.png)