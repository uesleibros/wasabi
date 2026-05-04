# Client‑Side Fragmentation (MTU‑Aware Sends)

Wasabi can split a large outgoing message into multiple WebSocket frames so that
each frame stays below the PMTU, avoiding IP fragmentation.

## When it is triggered

- Only when `AutoMTU = True` (default).
- The message size exceeds `mtu.OptimalFrameSize`.

## Algorithm (`WebSocketSendMTUAware`)

1. Convert the message to UTF‑8 bytes.
2. If size ≤ `OptimalFrameSize` → delegate to normal `WebSocketSend`.
3. Otherwise:
   - Send the first chunk as a TEXT frame with `FIN = 0`.
   - Send middle chunks as CONTINUATION frames, `FIN = 0`.
   - Send the last chunk as CONTINUATION, `FIN = 1`.

The same logic applies to binary messages via `WebSocketSendBinaryMTUAware`.
