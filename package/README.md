# Package

This folder contains **Wasabi.bas**, the single, self-contained file that brings a full networking stack to any VBA project. The module covers WebSocket (RFC 6455), raw TCP with optional TLS, and MQTT over WebSocket, all in one import with no external dependencies.

> When you import `Wasabi.bas`, the VBA runtime compiles it on the spot.
> There is no build step, no binary, and no packaging. What you see is exactly what runs inside the Office process.

## How to use

1. In the VBA editor, click **File -> Import File...**
2. Select `Wasabi.bas` from this folder.
3. No additional steps are required. No references, no tools, no setup.

After importing, you can call `WebSocketConnect` directly from any module.

> [!TIP]
> The complete API reference is available in [`docs/API_REFERENCE.md`](../docs/API_REFERENCE.md).

## What is included

Wasabi exposes three independent surface areas that share the same underlying transport and TLS engine.

**WebSocket** is the primary API. It handles the full RFC 6455 lifecycle: connection, framing, fragmentation, ping/pong, and graceful close. Connections support `permessage-deflate` compression, custom subprotocols, custom HTTP headers, automatic reconnection with exponential back-off, offline message queuing, and MTU-aware framing. Both text and binary frames are supported, including zero-copy receive for performance-sensitive paths.

**TCP** gives direct access to the raw socket layer, with or without TLS. This is useful when the remote endpoint speaks a line protocol, a proprietary binary format, or any framing scheme that is not WebSocket. The same TLS engine used by the WebSocket layer is available here, including client certificate authentication via thumbprint, subject name, or a PFX file loaded from disk.

**MQTT** runs on top of an established WebSocket connection and implements the MQTT 3.1.1 packet exchange, covering connect, publish (QoS 0/1/2), subscribe, unsubscribe, ping, and disconnect.

All three modes share a common set of transport features:

- TLS 1.2 and TLS 1.3 via the Windows SChannel/Schannel provider
- HTTP and SOCKS proxy support, with optional NTLM authentication and automatic proxy discovery from Internet Explorer settings
- IPv4 and IPv6 with configurable preference
- Configurable receive and inactivity timeouts
- Per-connection statistics and latency measurement
- Optional async mode via a hidden message window and `WSAAsyncSelect`
- A middleware and protocol-extension API for layering custom behaviour without modifying the module

## Connection handles

Every `Connect` call writes a handle to an `outHandle` variable. Pass that handle to any subsequent call to target a specific connection. If you have only one active connection, you can omit the handle argument and Wasabi will use the default. `WebSocketSetDefaultHandle` lets you change which connection is treated as the default.

## Error handling

Each connection stores the last error internally. Call `WebSocketGetLastError` to retrieve a typed `WasabiError` value, `WebSocketGetErrorDescription` for a human-readable message, and `WebSocketGetTechnicalDetails` for the low-level WSA or SChannel code. An optional error dialog can be enabled per connection with `WebSocketSetErrorDialog`.
