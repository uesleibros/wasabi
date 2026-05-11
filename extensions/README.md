# Wasabi Extensions

> [!IMPORTANT]
> The extension system is in the **design and validation phase**. The injection points already
> exist in the engine (`WasabiUseProtocol`, `WasabiUseMiddleware`, `WasabiUseCompression`),
> but the official stabilisation of the interfaces and their separation into distributable
> packages are part of the upcoming **Framework Era** milestone.

This directory contains blueprints, specifications, and reference implementations for
extensions: pluggable components that add high-level behaviour without modifying the
core `Wasabi.bas` transport engine.

## How extensions work

Wasabi's architecture separates the raw TCP/TLS transport from application-layer logic.
Custom code is injected as an extension by implementing a VBA class with the expected
interface and registering it against a connection handle.

```vb
' Example: registering a custom protocol handler
Dim myProto As New MyProtocol
WasabiUseProtocol myProto, handle
```

Extensions receive full lifecycle callbacks and have direct access to the connection,
enabling arbitrary extensibility without forking the core module.

## Extension types

| Type | Registration | Primary purpose |
|---|---|---|
| **Protocol Handler** | `WasabiUseProtocol` | Add an application-layer protocol (MQTT 5, AMQP, Modbus TCP, etc.) |
| **Middleware** | `WasabiUseMiddleware` | Intercept raw byte streams for logging, encryption, or transformation |
| **Compression Handler** | `WasabiUseCompression` | Replace or customise per-frame compression (LZ4, Brotli, Zstd) |

Detailed specifications and tutorials:

- **[Protocol Extension Guide](protocols.md)** - Interface, lifecycle, and a complete MQTT 5 example
- **[Middleware Extension Guide](middlewares.md)** - Intercepting inbound and outbound data, chaining, and encryption
- **[Compression Extension Guide](compression.md)** - Implementing custom `Deflate` and `Inflate` providers

## Lifecycle callbacks

All extensions share a common set of lifecycle callbacks fired by the engine at well-defined moments.

**Fired for all extension types:**

`OnConnect(handle)` is called immediately after the WebSocket handshake (or TCP connect) succeeds.

`OnDisconnect(handle)` is called before the connection is fully torn down. The handler may still attempt a final transmit at this point.

**Fired for middleware only:**

`OnBeforeSend(handle, data())` receives every outbound byte array before framing.

`OnAfterReceive(handle, data())` receives every inbound byte array after deframing and decryption.

**Fired for protocol handlers only:**

`OnTextMessage(handle, message As String)` delivers already-parsed text frames.

`OnBinaryMessage(handle, data() As Byte)` delivers already-parsed binary frames.

**Required for compression handlers:**

`Deflate(data(), windowBits, contextTakeover)` returns a compressed `Byte()`.

`Inflate(data(), windowBits, contextTakeover)` returns a decompressed `Byte()`.

> [!TIP]
> The engine never invokes an extension on a mismatched connection type. A protocol handler
> registered on a raw TCP handle is silently ignored.

## Integration model

The core `Wasabi.bas` module will remain a monolithic, zero-dependency file.
Extensions are distributed as additional `.cls` or `.bas` files that you import alongside it.
No COM registration, no external references, and no build tools are required.

Planned milestones:

- [x] Stabilise callback signatures for all extension types.
- [x] Publish reference implementations (for example, `ExtWasabiZlib.cls` for `permessage-deflate`).
- [ ] Provide a registration mechanism for default global middleware applied to every new connection.
