# Exponential Backoff Algorithm

## Formula

`delay = ReconnectBaseDelayMs * 2^(attempt - 1)`


Where `attempt` is the current `ReconnectAttempts` (1‑based).

## Capping

If the computed delay exceeds `MAX_RECONNECT_DELAY_MS` (30,000 ms), it is
clamped to 30 s.

## Example (base delay = 1000 ms)

| Attempt | Delay |
|---------|-------|
| 1       | 1 s   |
| 2       | 2 s   |
| 3       | 4 s   |
| 4       | 8 s   |
| 5       | 16 s  |
| 6+      | 30 s  |

## Why exponential?

It prevents flooding a recovering server with immediate retries, following
widely‑accepted transient‑error handling patterns.

## Why 30 s cap?

To ensure the client eventually gives up or retries within a reasonable window
while still providing back‑pressure during extended outages.
