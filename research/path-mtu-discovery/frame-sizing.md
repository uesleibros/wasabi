# Optimal Frame Size Calculation

`CalculateOptimalFrameSize(handle)` computes how many application bytes
can fit in a WebSocket frame without exceeding the MTU.

## Formula

available = MTU

- ETHERNET_HEADER (14)
- IP_HEADER_MIN (20) or IPv6 (40)
- TCP_HEADER_MIN (20)
- TLS overhead (5 + cbHeader + cbTrailer) if TLS
- WEBSOCKET_HEADER_MAX (14)


If the result is less than 125 bytes, it is clamped to 125 (the minimum
useful WebSocket frame). If larger than 65535 (max 16‑bit payload), it
is clamped to 65535.

The result is stored in `mtu.OptimalFrameSize` and used by
`WebSocketSendMTUAware` to split large messages.
