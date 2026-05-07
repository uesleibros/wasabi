# Wasabi Extensions

> [!IMPORTANT]
> The extension system is in the **design & validation phase**. Although the injection points  
> already exist in the engine (`WasabiUseProtocol`, `WasabiUseMiddleware`, `WasabiUseCompression`),  
> the official stabilisation of the interfaces and the separation into pluggable packages are  
> part of the upcoming **Framework Era** milestone.

This directory contains blueprints, specifications, and reference implementations for
**extensions** – plug‑in components that add high‑level behaviour without modifying the
core `Wasabi.bas` transport engine.

## How Extensions Work

Wasabi’s architecture separates the raw TCP/TLS “dumb pipe” from application‑layer logic.
Custom code can be injected as an **extension** by implementing a clean interface (a VBA class)
and registering it against a connection handle.

```vb
' Example: registering a custom protocol handler
Dim myProto As New MyMqttProtocol
WasabiUseProtocol myProto, handle
```

Extensions receive full lifecycle callbacks (`OnConnect` / `OnDisconnect`) and have access
to the underlying connection, enabling infinite extensibility without forking.

## Available Extension Types

| Type                     | Registration function       | Primary purpose |
|--------------------------|-----------------------------|-----------------|
| **Protocol Handler**     | `WasabiUseProtocol`         | Add a new application‑layer protocol (MQTT 5, AMQP, ModbusTCP, etc.) |
| **Middleware**           | `WasabiUseMiddleware`       | Intercept raw byte streams for logging, encryption, or transformation |
| **Compression Handler**  | `WasabiUseCompression`      | Replace or customise per‑frame compression (LZ4, Brotli, Zstd) |

Detailed specification and tutorials:

- **[Protocol Extension Guide](protocols.md)** – Interface, lifecycle, and complete MQTT 5 example
- **[Middleware Extension Guide](middlewares.md)** – Intercepting inbound/outbound data, chaining, and encryption
- **[Compression Extension Guide](compression.md)** – Implementing custom `Deflate`/`Inflate` providers

## Lifecycle Guarantees

All extensions receive the same well‑defined callbacks:

* **`OnConnect(handle)`** – called immediately after the WebSocket handshake (or TCP connect) succeeds.
* **`OnDisconnect(handle)`** – called before the connection is fully torn down; the handler may still
  attempt a final transmit if necessary.

Middlewares additionally receive:

* **`OnBeforeSend(handle, data())`** – every byte array *before* framing.
* **`OnAfterReceive(handle, data())`** – every byte array *after* deframing / decryption.

Compression handlers must provide:

* **`Deflate(data(), windowBits, contextTakeover)` → `Byte()`**
* **`Inflate(data(), windowBits, contextTakeover)` → `Byte()`**

Protocol handlers receive already‑parsed WebSocket messages:

* **`OnTextMessage(handle, message As String)`**
* **`OnBinaryMessage(handle, data() As Byte)`**

> [!TIP]
> The engine **never** calls your extension on a wrong‑mode connection – a protocol handler
> registered on a TCP handle will simply be ignored.

## Integration Plan

The core `Wasabi.bas` module will remain a monolithic, zero‑dependency file.
Extensions are distributed as **additional `.cls` / `.bas` files** that you import alongside Wasabi.
No COM registration, no external references – just plain VBA.

Future milestones:

- [ ] Stabilise the callback signatures for all extension types.
- [ ] Publish reference extensions (e.g., `ExtWasabiZlib.cls` for `permessage‑deflate`).
- [ ] Provide a registration mechanism for **default global middleware** (applied to every new connection).
