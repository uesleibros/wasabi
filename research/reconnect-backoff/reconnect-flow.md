# Reconnect Flow

## Trigger

`TryReconnect(handle)` is called when any send/receive error sets
`Connected = False` and `AutoReconnect` is enabled.

## State preservation

Before cleaning up the broken connection, all configuration is saved:
`OriginalUrl`, `AutoReconnect`, `ReconnectMaxAttempts`, `ReconnectBaseDelayMs`,
`ReconnectAttempts`, `PingIntervalMs`, `ReceiveTimeoutMs`, `LogCallback`,
`EnableErrorDialog`, custom headers, proxy settings, deflate parameters, etc.

## Delay

The reconnect attempt is delayed by an exponentially increasing amount (see
`backoff-algorithm.md`). The delay loop uses `DoEvents` (not `Application.OnTime`)
to remain portable to non‑VBA hosts like VB6/twinBASIC.

## Re‑creation

- `CleanupHandle(handle)` fully frees the old socket and security handles.
- Winsock is re‑initialized if it was cleaned up.
- The old connection structure is reused (arrays are re‑dimensioned to the saved
  buffer sizes).
- `ConnectHandle(handle, savedUrl)` is called to perform a fresh connection.

## Outcome

- On success: `ReconnectAttempts` is reset to 0, a log entry is emitted.
- On failure: `ReconnectAttempts` is incremented; if it reaches
  `ReconnectMaxAttempts`, `AutoReconnect` is disabled.
