# Nagle’s Algorithm and WebSocket Latency

## Problem

Nagle’s algorithm delays sending small TCP segments until either an outstanding
ACK is received or enough data accumulates to fill a MSS. For a protocol that
sends many tiny messages (WebSocket pings, pongs, MQTT acknowledgements), this
adds noticeable latency – often up to 200 ms.

## Example

1. Client sends an MQTT `PUBACK` (4 bytes).
2. Nagle holds the segment waiting for more data.
3. The broker doesn’t receive the ACK and retransmits, causing unnecessary
   traffic and delay.

## Wasabi’s solution

`WebSocketSetNoDelay True` enables `TCP_NODELAY`, forcing immediate transmission
of every `send()` call. This is particularly important for MQTT over WebSocket
and real‑time applications.

To avoid sending many tiny packets in a tight loop, Wasabi also offers
`WebSocketSendBatch` / `WebSocketSendBatchBinary`, which accumulate frames in a
64 KB buffer and flush them in a single `send()`.
