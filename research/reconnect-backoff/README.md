# Reconnect & Exponential Backoff

Wasabi can automatically reconnect after an unexpected disconnection, using
exponential backoff with configurable limits.

## Reference

- `TryReconnect()`, `ConnectHandle()`, and associated public properties
  (`AutoReconnect`, `ReconnectMaxAttempts`, etc.) in `Wasabi.bas`.
