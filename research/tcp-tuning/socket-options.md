# Socket Options

`ApplySocketOptions(handle)` is called right after the TCP connection is
established, before any TLS or WebSocket handshake.

## `TCP_NODELAY` (disable Nagle)

- **Constant:** `TCP_NODELAY = 1`, level `IPPROTO_TCP_SOL`.
- **Enabled when:** `m_Connections(handle).NoDelay = True` (configurable via
  `WebSocketSetNoDelay`).
- **Effect:** Turns off Nagle’s algorithm. Without it, small WebSocket frames
  (pings, MQTT PUBACK, control frames) would be delayed up to 200 ms while the
  TCP stack waits for more data.
- **Mitigation:** When many small messages need to be sent, Wasabi provides batch
  send functions (`WebSocketSendBatch`) that pack multiple frames into a single
  `send()` call, reducing the downside of disabled Nagle.

## `SO_KEEPALIVE`

- **Constant:** `SO_KEEPALIVE = 8`, level `SOL_SOCKET`.
- **Always enabled.**
- **Effect:** The OS sends TCP keep‑alive probes after a long idle period (2
  hours by default on Windows). This helps detect half‑open connections across
  NATs/firewalls. Wasabi also implements application‑level keep‑alive via
  `PingIntervalMs`, which operates on much shorter timescales.

## `SO_RCVBUF` and `SO_SNDBUF`

- **Value:** Both are set to `BUFFER_SIZE` (262,144 bytes = 256 KB).
- **Effect:** Enlarges the kernel socket buffers beyond the default (~8 KB). This
  lets the TCP stack receive more data while VBA is busy processing other
  messages, avoiding dropped packets.
- **Trade‑off:** Slightly higher non‑paged kernel memory usage. 256 KB is
  negligible on modern systems.

## `FIONBIO` (non‑blocking mode)

- **Used only during connection establishment** inside `ResolveAndConnect` for
  the Happy Eyeballs sockets. Once the winning socket is chosen, it is switched
  back to blocking mode (`FIONBIO = 0`).
- **Why:** Allows non‑blocking `connect()` and `select()` to race IPv6 and IPv4
  without freezing the VBA host.
