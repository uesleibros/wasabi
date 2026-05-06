# Wasabi Extensions (Planning Phase)

> [!IMPORTANT]
> This entire directory and the extension system are currently in the ![](../resources/svg/planning.svg) **conceptual stage**. The current production release of Wasabi (v2.x) remains a monolithic `.bas` module. Do not attempt to use these features in production yet.

This directory is the blueprint for the future **Wasabi Extension Ecosystem**. We are moving toward a modular framework where the community can plug in custom logic without modifying the core WinAPI transport engine.

### The Roadmap Vision

In this upcoming era, `Wasabi.bas` will serve as the high-performance **Dumb Pipe**, handling only raw Sockets, Schannel (TLS), and memory management. Everything else becomes an injectable "Lego piece."

#### 1. Custom Protocols (`/protocols`)
While Wasabi ships with "batteries included" for **WebSocket** and **MQTT 3.1.1**, this architecture will allow the community to implement:
*   **MQTT 5.0**: Advanced support for user properties and session management.
*   **Socket.IO**: A translator for Engine.IO packets (types 0, 2, 3, 42) to bridge VBA with modern web servers.
*   **ModbusTCP**: Direct communication for Industrial IoT and PLC integration.
*   **AMQP**: Connectivity for enterprise message brokers like RabbitMQ.

#### 2. Middleware Pipeline (`/middlewares`)
Inspired by the **Express.js** pipeline, allowing you to intercept data before it hits the wire or before it reaches your VBA code:
*   **Security Interceptors**: Automatically inject OAuth2 tokens, JWTs, or AWS Signature V4 headers.
*   **Audit Logging**: Real-time telemetry for packet sizes and RTT latency without bloat.
*   **End-to-End Encryption**: Plug in AES-256 or custom encryption layers inside the secure tunnel.

#### 3. Modular Compression (`/compression`)
> [!NOTE]
> We plan to decouple the `zlib1.dll` dependency from the main module.
*   **Core-Only**: Keep your project lean if you don't need compression.
*   **Wasabi-Zlib**: Simply import the `ExtWasabiZlib.cls` extension to re-enable `permessage-deflate` support.

### Planned Architecture
The system will utilize **Late Binding** (`Object`) and `Collections` to maintain the ease of a `.bas` file while gaining infinite extensibility.

> [!TIP]
> **Conceptual usage example:**
> ```vb
> ' Future implementation concept:
> Dim sIO As Object
> Set sIO = New WasabiSocketIO
>
> Wasabi.UseProtocol sIO, handle
> Wasabi.WebSocketConnect "wss://server.com/socket.io/?EIO=4"
> ```

> [!IMPORTANT]
> If you are an experienced network engineer or VBA architect and wish to contribute to the **Extension API Specification**, please open a Pull Request/Issue or check the Roadmap.
