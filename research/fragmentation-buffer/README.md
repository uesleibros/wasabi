# Fragmentation Buffer

WebSocket allows messages to be split across multiple frames. Wasabi supports
reassembling received fragments and (via MTU‑aware sends) generating fragmented
outgoing messages.

## Reference

- `ProcessFrames()` and `WebSocketSendMTUAware()` in `Wasabi.bas`.
