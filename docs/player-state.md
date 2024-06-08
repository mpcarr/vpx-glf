# Player State

The player state stores variables per player. 

## Set Player State

```SetPlayerState``` is used to put data into the player state. You need a key and value.
```
SetPlayerState "current_ball", 1
```

## Get Player State

```GetPlayerState``` is used to get data out of the player state.

```
GetPlayerState "current_ball"
```

## Player Events

When a player state value changes it will post an event so you can respond to that change. e.g. light changes, dmd updates.

### Add Player State Event Listener

To start monitoring a player event.

The parameters are: 

- Player State Key: e.g. "score", this is the value you want to monitor
- Key: A unqiue key for this listnener. You can have multiple subs monitoring the same value doing different things. This key is used to separate those.
- Callback: The Sub or Function to call when the value changes
- Priority: callbacks will be called in priority order, the higher priorities will be called first.
- Args: Any arguments you want to pass along to the callback.

```
AddPlayerStateEventListener "score", "player_score_changed", "UpdateDMD", 1000, Null
```

### Remove Player State Event Listener

To stop monitoring a player event.

```
RemovePlayerStateEventListener "score", "player_score_change"
```