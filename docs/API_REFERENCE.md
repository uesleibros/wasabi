# API Reference

This document describes the complete public interface of **Wasabi.bas**, a native WebSocket, WSS, and TCP client for VBA built on raw Winsock and Schannel. It implements the WebSocket protocol as defined by RFC 6455 and handles TLS through the Windows SSPI Schannel provider.

## Table of Contents

- [Core Concepts](#core-concepts)
  - [Handles and the Connection Pool](#handles-and-the-connection-pool)
  - [The Default Handle](#the-default-handle)
  - [Polling Model](#polling-model)
  - [Async Model](#async-model)
  - [Message Queues and Capacity](#message-queues-and-capacity)
  - [Connection Modes](#connection-modes)
  - [Extension Architecture](#extensions-plug-in-architecture)
- [Connection Management](#connection-management)
- [Sending Data](#sending-data)
- [Receiving Data](#receiving-data)
- [Queue Inspection and Control](#queue-inspection-and-control)
- [Control Frames and Heartbeat](#control-frames-and-heartbeat)
- [Reconnect and Reliability](#reconnect-and-reliability)
- [Buffer and Performance Configuration](#buffer-and-performance-configuration)
- [MTU Management](#mtu-management)
- [Proxy Configuration](#proxy-configuration)
- [Handshake Customization](#handshake-customization)
- [Timeout Configuration](#timeout-configuration)
- [Diagnostics and Monitoring](#diagnostics-and-monitoring)
- [Logging and User Feedback](#logging-and-user-feedback)
- [Security Configuration](#security-configuration)
- [Extension System](#extension-system-pluggable-middleware-protocol-and-compression)
- [TCP Client](#tcp-client)
- [MQTT Client](#mqtt-client)
- [Compression](#compression-and-extensions)
- [Error Reference](#error-reference)
- [Practical Patterns](#practical-patterns)
- [Operational Caveats](#operational-caveats)

## Core Concepts

### Handles and the Connection Pool

Wasabi maintains an internal pool of up to 64 simultaneous connections. Each connection is identified by an integer handle, which is an index into this pool. Handles are allocated automatically when you call a connect function and are released when you disconnect.

Most public functions accept an optional `handle` parameter of type `Long`. Valid handles range from `0` to `63` inclusive.

> [!IMPORTANT]
> Do not assume that handle `0` is invalid or uninitialized. `0` is a fully valid connection handle.

The recommended pattern for multi-connection applications is to always declare, store, and pass handles explicitly.

```vb
Dim h As Long
Dim ok As Boolean

ok = WebSocketConnect("wss://echo.websocket.org", h)

If ok Then
    WebSocketSend "Message", h
    Debug.Print WebSocketReceive(h)
    WebSocketDisconnect h
End If
```

### The Default Handle

Wasabi tracks one global default handle. When you call a function without providing a handle, Wasabi resolves it to the current default. Every successful call to `WebSocketConnect` automatically promotes the new connection to the default.

> [!TIP]
> Single-connection workbooks can safely omit the handle parameter in all calls, since the default handle is always up to date.

```vb
WebSocketConnect "wss://echo.websocket.org"
WebSocketSend "Hello"
Debug.Print WebSocketReceive()
WebSocketDisconnect
```

> [!WARNING]
> In multi-connection scenarios, omitting the handle may silently route calls to the wrong connection. Always pass handles explicitly when managing more than one connection.

### Polling Model

VBA is single-threaded. There is one execution thread shared between your code, the Office runtime, and all Wasabi operations. Background threads are not available in the VBA runtime.

Wasabi cannot push messages to your code automatically. Received data accumulates in internal buffers and is processed only when you call one of the receive functions.

Each call to `WebSocketReceive` or `WebSocketReceiveAll` triggers the following internal sequence:

1. Calls internal maintenance to check ping intervals and inactivity timeouts.
2. Processes any already-buffered decoded frames.
3. Checks the OS socket buffer for available data using `ioctlsocket` with `FIONREAD`.
4. If data is available, reads it with `recv`.
5. If TLS is active, passes raw bytes through `DecryptMessage`.
6. Parses WebSocket frames and routes each by opcode.
7. Enqueues complete text and binary messages into their respective ring buffers.
8. Returns the oldest queued message.

> [!IMPORTANT]
> Your code must call a receive function regularly for any of the following to work: message delivery, automatic ping scheduling, inactivity timeout detection, and auto reconnect triggering.

### Async Model

Wasabi provides an event-driven async model as an alternative to the polling model. Instead of calling receive functions in a loop, you register a handler object and let the Windows message pump deliver socket events to your callbacks automatically.

The async system is implemented via `WSAAsyncSelect`, which instructs Windows to post socket event notifications (`FD_READ`, `FD_WRITE`, `FD_CLOSE`, `FD_CONNECT`) as messages to a hidden Win32 window created internally by Wasabi. A native machine-code thunk subclasses that window and dispatches each event to your handler object using `CallByName`.

This means socket callbacks fire naturally while the Excel runtime is idle, without any polling loop on your part. If your code is actively running a tight loop without `DoEvents`, incoming messages will queue in the Windows message pump and be dispatched as soon as your code finishes or yields.

> [!IMPORTANT]
> The async handler object must be stored in a module-level or workbook-level variable. If it is declared as a local variable inside a `Sub`, it will be destroyed when that `Sub` returns, and subsequent callbacks will have no target.

```vb
' In a standard module at module level
Private g_Handler As cWasabiAsync
Private g_Handle As Long

Public Sub StartAsync()
    Set g_Handler = New cWasabiAsync
    If WebSocketConnect("wss://stream.binance.com:9443/ws/btcusdt@trade", g_Handle) Then
        WasabiUseAsync g_Handler, g_Handle
    End If
    ' Sub ends here. Callbacks fire while Excel is idle.
End Sub
```

The handler class must implement the following five methods:

```vb
Public Sub OnConnect(ByVal handle As Long)
Public Sub OnReceive(ByVal handle As Long)
Public Sub OnReadyToSend(ByVal handle As Long)
Public Sub OnClose(ByVal handle As Long)
Public Sub OnError(ByVal handle As Long, ByVal errorCode As Long, ByVal eventType As Long)
```

`OnReceive` is called whenever data arrives on the socket. Inside it, drain the queue normally using `WebSocketReceive` or `WebSocketGetPendingCount`.

`OnReadyToSend` is called when the socket becomes writable after a previous send would have blocked. This is useful for backpressure scenarios and can be left empty if not needed.

`OnError` receives the raw WSA error code and the event type that triggered the error (`FD_READ`, `FD_CONNECT`, etc.).

A complete handler class example:

```vb
Option Explicit

Public Sub OnConnect(ByVal handle As Long)
    Debug.Print "Connected. Handle: " & handle
End Sub

Public Sub OnReceive(ByVal handle As Long)
    Dim msg As String
    Do While WebSocketGetPendingCount(handle) > 0
        msg = WebSocketReceive(handle)
        Debug.Print "Received: " & msg
    Loop
End Sub

Public Sub OnReadyToSend(ByVal handle As Long)
End Sub

Public Sub OnClose(ByVal handle As Long)
    Debug.Print "Connection closed. Handle: " & handle
End Sub

Public Sub OnError(ByVal handle As Long, ByVal errorCode As Long, ByVal eventType As Long)
    Debug.Print "Async error on handle " & handle & " | Code: " & errorCode & " | Event: " & eventType
End Sub
```

> [!NOTE]
> The async and polling models are not mutually exclusive. Registering an async handler does not prevent you from calling `WebSocketReceive` manually. Both paths share the same internal message queue.

> [!WARNING]
> Always call `WebSocketDisconnectAll` or `WebSocketDisconnect` before resetting the VBA project or closing the workbook. The async system uses a native thunk that holds a pointer to the VBA runtime. Resetting without cleanup can crash the host application. The thunk contains a guard that checks whether the VBA runtime is still alive before dispatching, but explicit cleanup is the safe practice.

### Message Queues and Capacity

Wasabi uses two separate circular queues per connection: one for text messages and one for binary messages. Each has a fixed capacity of 512 entries.

When a queue reaches capacity, new incoming messages of that type are discarded and a log warning is emitted.

> [!WARNING]
> If your application receives messages at a high rate, ensure your polling loop drains the queue frequently enough to avoid drops.

You can check queue state at any time.

```vb
Dim pending As Long
Dim capacity As Long

pending = WebSocketGetPendingCount(h)
capacity = WebSocketGetQueueCapacity(h)

Debug.Print "Pending text messages:", pending
Debug.Print "Remaining capacity:", capacity
```

### Connection Modes

Every handle operates in one of three modes, represented by the `WasabiConnectionMode` enum:

| Value | Constant | Description |
|:---|:---|:---|
| 0 | `MODE_WEBSOCKET` | Standard WebSocket or WSS connection |
| 1 | `MODE_TCP` | Plain TCP connection |
| 2 | `MODE_TCP_TLS` | TCP connection with TLS 1.2/1.3 |

WebSocket and TCP handles share the same pool, the same proxy infrastructure, the same TLS stack, and the same MTU discovery engine. The mode is set automatically by the connect function used and cannot be changed after connection.

WebSocket-specific functions (`WebSocketSend`, `WebSocketReceive`, MQTT, etc.) will silently exit if called on a TCP handle. TCP-specific functions (`TcpSend`, `TcpReceive`, etc.) will silently exit if called on a WebSocket handle.

### Extensions: Plug-in Architecture

Wasabi provides three dedicated extension points that allow you to inject custom behaviour without modifying the core module:

**Protocol Handler** (`WasabiUseProtocol`): intercepts parsed WebSocket text and binary messages, useful for implementing application-layer protocols like MQTT 5 or custom binary parsers.

**Middleware** (`WasabiUseMiddleware`): intercepts raw byte arrays before they are framed (outbound) or after they are deframed (inbound). Use for logging, encryption, or header injection.

**Compression Handler** (`WasabiUseCompression`): replaces the built-in `permessage-deflate` path with any algorithm (LZ4, Brotli, Zstd). The core is completely independent of any compression library.

Extensions are described in detail in the [Extension System](#extension-system-pluggable-middleware-protocol-and-compression) section and in the [extensions/](extensions/) directory.

## Connection Management

### WebSocketConnect

```vb
Public Function WebSocketConnect(ByVal url As String, Optional ByRef outHandle As Long = -1, Optional ByVal DeflateEnabled As Boolean = False, Optional ByVal DeflateContextTakeover As Boolean = True, Optional ByVal SubProtocol As String = "") As Boolean
```

Opens a new WebSocket connection to the specified URL.

This function executes the complete connection sequence:

1. Initializes Winsock if not already active.
2. Allocates a slot in the connection pool.
3. Parses the URL into host, port, path, and scheme.
4. Resolves the hostname via `gethostbyname` or `inet_addr` if a literal IP is supplied.
5. Creates a non-blocking TCP socket.
6. Initiates connection with `connect` and waits for completion using `select`.
7. Applies `TCP_NODELAY`, `SO_KEEPALIVE`, `SO_RCVBUF`, and `SO_SNDBUF` socket options.
8. If proxy is configured, sends an HTTP CONNECT (or SOCKS5) request and verifies the tunnel.
9. If the scheme is `wss://`, performs a TLS handshake via Schannel SSPI.
10. If configured, validates the server certificate chain.
11. Queries `SecPkgContext_StreamSizes` to determine TLS record framing parameters.
12. Sends the HTTP/1.1 WebSocket upgrade request.
13. Validates the `Sec-WebSocket-Accept` header value.
14. Marks the connection state as `STATE_OPEN` and records the connection timestamp.

**Parameters:**

`url` — A WebSocket URL beginning with `ws://` or `wss://`. Custom ports may be specified inline, for example `wss://api.example.com:8443/ws`.

`outHandle` — Receives the allocated integer handle. Set to `-1` on failure.

`DeflateEnabled` — Set to `True` to request `permessage-deflate` during the handshake. You must also register a compression handler via `WasabiUseCompression` for actual compression to occur.

`DeflateContextTakeover` — When `True`, the compression context is reused across messages for better ratios.

`SubProtocol` — Optional string for the `Sec-WebSocket-Protocol` header, for example `"mqtt"`.

**Returns:** `True` if the connection was fully established. `False` if any step in the sequence failed.

```vb
Dim h As Long

If WebSocketConnect("wss://echo.websocket.org", h) Then
    Debug.Print "Connected on handle", h
Else
    Debug.Print "Error code:", WebSocketGetLastError()
    Debug.Print "System code:", WebSocketGetLastErrorCode()
    Debug.Print "Details:", WebSocketGetTechnicalDetails()
End If
```

> [!IMPORTANT]
> Configure all options (proxy, custom headers, buffer sizes, MTU, no-delay, extensions) before calling `WebSocketConnect`. These settings are applied during the connection sequence and cannot be changed mid-connection.

> [!CAUTION]
> The connection pool has a hard limit of 64 simultaneous connections. Attempting to connect beyond this limit returns `ERR_MAX_CONNECTIONS`.

### WebSocketDisconnect

```vb
Public Sub WebSocketDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Gracefully terminates a connection. This function disables auto reconnect for the handle, sends a WebSocket Close frame with status code 1000 if the connection is active, closes the underlying TCP socket, releases the Schannel security context and credential handle, and clears all internal connection state.

If no other connections remain active after this call, Wasabi also calls `WSACleanup` to release Winsock resources.

```vb
WebSocketDisconnect h

If WebSocketGetConnectionCount() = 0 Then
    Debug.Print "All connections closed"
End If
```

> [!NOTE]
> Calling `WebSocketDisconnect` always disables auto reconnect for the target handle. Disconnecting means you chose to close the connection, so auto reconnect would be unexpected behaviour.

### WebSocketDisconnectAll

```vb
Public Sub WebSocketDisconnectAll()
```

Iterates through the entire pool and disconnects every active connection. Also shuts down the async window and thunk if they were initialized.

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    WebSocketDisconnectAll
End Sub
```

### WebSocketIsConnected

```vb
Public Function WebSocketIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns the current connected state of the handle. This is a lightweight state check that evaluates whether the internal state is `STATE_OPEN` without touching the socket.

> [!NOTE]
> This function reports the last known state. It does not actively probe the socket. A connection may have dropped between the last receive call and this check.

```vb
Do While WebSocketIsConnected(h)
    Dim msg As String
    msg = WebSocketReceive(h)
    If msg <> "" Then Debug.Print msg
    DoEvents
Loop
```

### WebSocketSendClose

```vb
Public Function WebSocketSendClose(Optional ByVal code As Integer = 1000, Optional ByVal reason As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a WebSocket Close control frame and transitions the connection to `STATE_CLOSING`. The status code occupies the first two bytes of the Close frame payload in big-endian byte order, followed by the UTF-8 encoded reason string.

Standard close codes from RFC 6455:

| Code | Meaning |
|:---|:---|
| 1000 | Normal closure |
| 1001 | Endpoint going away |
| 1002 | Protocol error |
| 1003 | Unsupported data |
| 1007 | Invalid frame payload data |
| 1008 | Policy violation |
| 1009 | Message too large |
| 1011 | Internal server error |

```vb
WebSocketSendClose 1001, "Client shutting down", h
```

> [!NOTE]
> `WebSocketDisconnect` calls `WebSocketSendClose` internally. Call `WebSocketSendClose` directly only if you need to specify a custom close code or reason before cleanup.

### WasabiUseAsync

```vb
Public Sub WasabiUseAsync(ByVal handler As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers an async event handler for the specified connection and switches it to event-driven mode. After this call, socket events are dispatched to the handler object via the Windows message pump instead of requiring manual polling.

`handler` must be an object implementing all five async callback methods described in the [Async Model](#async-model) section.

Calling this function on an already-connected handle immediately registers the socket with `WSAAsyncSelect`, so events begin arriving on the next idle cycle.

```vb
Dim g_Handler As cWasabiAsync
Dim g_Handle As Long

Set g_Handler = New cWasabiAsync
WebSocketConnect "wss://stream.binance.com:9443/ws/btcusdt@trade", g_Handle
WasabiUseAsync g_Handler, g_Handle
```

> [!WARNING]
> The handler object must remain alive for as long as the connection is active. Store it at module or workbook level, never as a local variable.

## Sending Data

### WebSocketSend

```vb
Public Function WebSocketSend(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a text message. Internally, Wasabi converts the VBA Unicode string to UTF-8 using `WideCharToMultiByte`, then constructs a masked WebSocket text frame. The payload length is encoded using the appropriate tier:

- 7-bit for payloads up to 125 bytes.
- 16-bit with prefix `0x7E` for payloads up to 65535 bytes.
- 64-bit with prefix `0x7F` for larger payloads.

The frame is transmitted via `send` in a loop to handle partial writes. If a compression handler is registered and `permessage-deflate` was negotiated, the payload is compressed before framing.

**Returns:** `True` if all frame bytes were successfully written to the socket, or if offline queueing is enabled and the message was successfully buffered.

```vb
If Not WebSocketSend("Hello from VBA", h) Then
    Debug.Print "Send failed:", WebSocketGetTechnicalDetails(h)
End If
```

> [!NOTE]
> If the connection drops and `WebSocketSetOfflineQueueing` is enabled, messages are queued in memory and will be transmitted automatically once `AutoReconnect` restores the connection. Otherwise, sending on a disconnected handle returns `False` and logs `ERR_NOT_CONNECTED`.

### WebSocketSendBinary

```vb
Public Function WebSocketSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a binary message using WebSocket opcode `0x02`. The byte array is wrapped in a masked binary frame using the same length encoding tiers as text messages.

```vb
Dim payload(0 To 7) As Byte
payload(0) = &H01 : payload(1) = &H02 : payload(2) = &HFF : payload(3) = &H00
payload(4) = &HAB : payload(5) = &HCD : payload(6) = &HEF : payload(7) = &H42

If WebSocketSendBinary(payload, h) Then
    Debug.Print "Binary sent"
End If
```

### WebSocketBroadcast

```vb
Public Function WebSocketBroadcast(ByVal message As String) As Long
```

Sends the same text message to every active connection in the pool. Returns the count of connections that received the message successfully.

### WebSocketBroadcastBinary

```vb
Public Function WebSocketBroadcastBinary(ByRef data() As Byte) As Long
```

Sends the same binary payload to every active connection.

### Advanced Sending

#### WebSocketSendBatch

```vb
Public Function WebSocketSendBatch(ByRef messages() As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends multiple text messages in a single TCP write (or a minimal number of writes) to reduce system call overhead. All messages are packed into a contiguous byte buffer before transmission.

#### WebSocketSendBatchBinary

```vb
Public Function WebSocketSendBatchBinary(ByRef messages() As Variant, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends multiple binary payloads in a single TCP write. Each element of the `messages` array must be a `Byte()` array.

#### WebSocketSendMTUAware

```vb
Public Function WebSocketSendMTUAware(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a text message using WebSocket fragmentation automatically tuned to the connection's current MTU. If the message fits in one frame, it behaves identically to `WebSocketSend`. Otherwise, the message is split into fragments sized to `OptimalFrameSize` and sent as a sequence of continuation frames.

#### WebSocketSendBinaryMTUAware

```vb
Public Function WebSocketSendBinaryMTUAware(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Like `WebSocketSendMTUAware`, but for binary payloads (opcode `0x02` with continuation fragments).

## Receiving Data

### WebSocketReceive

```vb
Public Function WebSocketReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the oldest queued text message. Returns an empty string if none is available.

Each invocation drives the full internal polling and maintenance cycle described in the Core Concepts section. If a protocol handler is registered, text messages are delivered to its `OnTextMessage` callback instead of being queued; `WebSocketReceive` will then always return an empty string.

```vb
Dim msg As String
msg = WebSocketReceive(h)
If msg <> "" Then
    Debug.Print "Got:", msg
End If
```

### WebSocketReceiveAll

```vb
Public Function WebSocketReceiveAll(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String()
```

Drains the complete text queue and returns all pending messages as a string array. This is more efficient than calling `WebSocketReceive` in a loop when you expect bursts of messages.

### WebSocketReceiveBinary

```vb
Public Function WebSocketReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
```

Returns the oldest binary message from the queue. If a protocol handler is registered, binary messages are sent to `OnBinaryMessage` instead, and this function returns an empty array.

### WebSocketReceiveBinaryCheck

```vb
Public Function WebSocketReceiveBinaryCheck(ByRef outData() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Populates `outData` with the oldest binary message and returns `True` if data was available. Preferred when you want explicit control flow.

### Zero-Copy Receiving

Zero-copy functions allow direct access to internal data buffers without creating copies. They must be enabled per-connection via `WebSocketSetZeroCopy`.

#### WebSocketReceiveZeroCopy

```vb
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns a pointer to the internal string buffer of the next text message, along with its length in characters.

#### WebSocketReceiveBinaryZeroCopy

```vb
Public Function WebSocketReceiveBinaryZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns a pointer to the internal byte array of the next binary message, along with its length in bytes.

## Queue Inspection and Control

### WebSocketGetPendingCount

```vb
Public Function WebSocketGetPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the number of text messages currently in the queue.

### WebSocketGetBinaryPendingCount

```vb
Public Function WebSocketGetBinaryPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the number of binary messages currently in the queue.

### WebSocketGetQueueCapacity

```vb
Public Function WebSocketGetQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns how many more text messages can be queued before overflow begins.

### WebSocketGetBinaryQueueCapacity

```vb
Public Function WebSocketGetBinaryQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns how many more binary messages can be queued before overflow begins.

### WebSocketPeek

```vb
Public Function WebSocketPeek(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the next queued text message without removing it from the queue.

### WebSocketFlushQueue

```vb
Public Sub WebSocketFlushQueue(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Immediately clears all queued text and binary messages for the specified connection.

> [!WARNING]
> This operation is irreversible. All queued messages are discarded immediately. Data still pending in the OS socket buffer is not affected and may appear in subsequent receive calls.

## Control Frames and Heartbeat

### WebSocketSendPing

```vb
Public Function WebSocketSendPing(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a Ping control frame with an optional text payload up to 125 bytes.

### WebSocketSendPong

```vb
Public Function WebSocketSendPong(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a Pong control frame with an optional payload. Wasabi automatically responds to incoming server Pings with a matching Pong during frame processing.

### WebSocketSetPingInterval

```vb
Public Sub WebSocketSetPingInterval(ByVal intervalMs As Long, Optional ByVal jitterMaxMs As Long = 0, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the base interval in milliseconds for automatic Ping frames. Set `intervalMs` to `0` to disable.

`jitterMaxMs` optionally adds a pseudo-random variance to the ping interval. For example, setting `intervalMs` to `30000` and `jitterMaxMs` to `5000` means pings will fire randomly between 30 and 35 seconds. This is important for bypassing strict anti-bot gateway filters that drop clients with perfectly timed heartbeats.

```vb
WebSocketSetPingInterval 30000, 5000, h
```

> [!NOTE]
> In async mode, automatic pings are driven by the `OnReceive` callback path, which internally calls `TickMaintenance`. If no data arrives for a long period, pings may not fire on schedule. For reliable heartbeating in async mode, consider calling `WebSocketSendPing` from a timer or `Application.OnTime` routine.

### WebSocketGetLatency

```vb
Public Function WebSocketGetLatency(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the most recent round-trip time (RTT) in milliseconds, measured from the last Ping frame sent to the corresponding Pong frame received.

## Reconnect and Reliability

### WebSocketSetAutoReconnect

```vb
Public Sub WebSocketSetAutoReconnect(ByVal enabled As Boolean, Optional ByVal maxAttempts As Long = DEFAULT_RECONNECT_MAX_ATTEMPTS, Optional ByVal baseDelayMs As Long = DEFAULT_RECONNECT_BASE_DELAY_MS, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables automatic reconnection using exponential backoff.

When a session loss is detected during polling, Wasabi saves all connection settings, cleans up resources, waits for the calculated delay, and invokes `ConnectHandle` with the original URL to re-establish the session. The delay between attempts doubles on each failure (Attempt 1: `baseDelayMs`, Attempt 2: `baseDelayMs * 2`, etc.), capped at 30 seconds.

### WebSocketSetOfflineQueueing

```vb
Public Sub WebSocketSetOfflineQueueing(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables the offline retention queue. When enabled, any calls to `WebSocketSend`, `WebSocketSendBinary`, or `MqttPublish` while the socket is disconnected will be stored in an internal buffer rather than discarded. Once `AutoReconnect` successfully re-establishes the connection, all buffered messages are automatically flushed in the exact order they were queued.

```vb
WebSocketSetOfflineQueueing True, h
```

### WebSocketGetReconnectInfo

```vb
Public Function WebSocketGetReconnectInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the current reconnect configuration and attempt count as a pipe-delimited string.

## Buffer and Performance Configuration

### WebSocketSetBufferSizes

```vb
Public Sub WebSocketSetBufferSizes(ByVal bufferSize As Long, ByVal fragmentSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Overrides the default sizes for the receive buffer and the fragment reassembly buffer. Default size for both is 256 KB (262144 bytes).

### WebSocketSetNoDelay

```vb
Public Function WebSocketSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Controls the `TCP_NODELAY` socket option, which disables Nagle's algorithm when enabled, reducing latency for small, frequent packets.

## MTU Management

Wasabi can automatically detect the path MTU and adjust WebSocket frame fragmentation to avoid IP fragmentation.

### WebSocketSetMTU

```vb
Public Sub WebSocketSetMTU(ByVal mtu As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the static MTU value for frame sizing calculations. Valid range: 576 to 9000. Default is 1500.

### WebSocketGetMTU

```vb
Public Function WebSocketGetMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the current MTU value used for frame sizing.

### WebSocketSetAutoMTU

```vb
Public Sub WebSocketSetAutoMTU(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables automatic MTU discovery based on the TCP MSS.

### WebSocketGetOptimalFrameSize

```vb
Public Function WebSocketGetOptimalFrameSize(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the calculated optimal frame payload size in bytes for the current MTU and TLS overhead configuration.

### WebSocketGetMTUInfo

```vb
Public Function WebSocketGetMTUInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited summary of the current MTU state.

### WebSocketProbeMTU

```vb
Public Sub WebSocketProbeMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Forces an immediate MTU probe on the active socket.

## Proxy Configuration

### WebSocketAutoDiscoverProxy

```vb
Public Sub WebSocketAutoDiscoverProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Automatically discovers and applies the Windows system proxy settings (including PAC scripts) for the current user.

### WebSocketSetProxy

```vb
Public Sub WebSocketSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal proxyType As Long = PROXY_TYPE_HTTP, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Configures an HTTP CONNECT or SOCKS5 proxy for the connection.

### WebSocketSetProxyNtlm

```vb
Public Sub WebSocketSetProxyNtlm(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables NTLM/Kerberos authentication for HTTP proxies using the currently logged-on Windows user's credentials.

### WebSocketClearProxy

```vb
Public Sub WebSocketClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes all proxy settings for the specified connection handle.

### WebSocketGetProxyInfo

```vb
Public Function WebSocketGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a summary of the current proxy configuration.

## Handshake Customization

### WebSocketAddHeader

```vb
Public Sub WebSocketAddHeader(ByVal headerName As String, ByVal headerValue As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Adds a custom HTTP header to the WebSocket upgrade request.

### WebSocketClearHeaders

```vb
Public Sub WebSocketClearHeaders(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes all custom headers added via `WebSocketAddHeader`.

### WebSocketSetSubProtocol

```vb
Public Sub WebSocketSetSubProtocol(ByVal protocol As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the value sent in the `Sec-WebSocket-Protocol` header during the handshake.

### WebSocketGetSubProtocol

```vb
Public Function WebSocketGetSubProtocol(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the configured subprotocol string.

## Timeout Configuration

### WebSocketSetReceiveTimeout

```vb
Public Sub WebSocketSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the timeout in milliseconds for internal `select()` calls used when waiting for data during handshake and TLS operations.

### WebSocketSetInactivityTimeout

```vb
Public Sub WebSocketSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the maximum allowed period without receiving any data from the server. If this duration elapses, the connection is treated as stale and closed.

## Diagnostics and Monitoring

### WebSocketGetLastError

```vb
Public Function WebSocketGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
```

Returns the most recent `WasabiError` enumeration value for the connection.

### WebSocketGetLastErrorCode

```vb
Public Function WebSocketGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the most recent native system error code (WSA or SSPI).

### WebSocketGetTechnicalDetails

```vb
Public Function WebSocketGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a detailed technical description of the most recent error.

### WebSocketGetErrorDescription

```vb
Public Function WebSocketGetErrorDescription(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a single human-readable string combining the error category, system code, and technical details.

### WebSocketGetStats

```vb
Public Function WebSocketGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a snapshot of connection metrics as a pipe-delimited string.

### WebSocketResetStats

```vb
Public Sub WebSocketResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Resets all counters to zero and updates the connected timestamp.

### WebSocketGetUptime

```vb
Public Function WebSocketGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns how many seconds the connection has been active.

### WebSocketGetHost

```vb
Public Function WebSocketGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the hostname resolved during connection.

### WebSocketGetPort

```vb
Public Function WebSocketGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the port used during connection.

### WebSocketGetPath

```vb
Public Function WebSocketGetPath(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the path component of the original URL.

### WebSocketGetConnectionCount

```vb
Public Function WebSocketGetConnectionCount() As Long
```

Returns the total number of currently active connections.

### WebSocketGetAllHandles

```vb
Public Function WebSocketGetAllHandles() As Long()
```

Returns an array of all currently active handles.

### WebSocketGetCloseCode

```vb
Public Function WebSocketGetCloseCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Integer
```

Returns the close code from the last Close frame received from the server.

### WebSocketGetCloseReason

```vb
Public Function WebSocketGetCloseReason(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the close reason string from the last Close frame.

### WebSocketGetCloseInfo

```vb
Public Function WebSocketGetCloseInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited summary of the last close event.

## Logging and User Feedback

### WebSocketSetLogCallback

```vb
Public Sub WebSocketSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a VBA macro as a log receiver. Wasabi calls this macro using `Application.Run` whenever it emits an internal diagnostic message.

### WebSocketSetErrorDialog

```vb
Public Sub WebSocketSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Controls whether connection errors trigger a `MsgBox` dialog. Disabled by default.

## Security Configuration

### WebSocketSetCertValidation

```vb
Public Sub WebSocketSetCertValidation(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables server certificate chain validation after the TLS handshake. Disabled by default.

### WebSocketSetRevocationCheck

```vb
Public Sub WebSocketSetRevocationCheck(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables CRL/OCSP revocation checking during server certificate validation.

### WebSocketSetClientCert

```vb
Public Sub WebSocketSetClientCert(ByVal thumbprintOrSubject As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Configures a client certificate for TLS mutual authentication (mTLS) from the Windows Personal store.

### WebSocketSetClientCertPfx

```vb
Public Sub WebSocketSetClientCertPfx(ByVal pfxPath As String, ByVal pfxPassword As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Loads a client certificate from a PFX file for TLS mutual authentication.

### WebSocketSetHttp2

```vb
Public Sub WebSocketSetHttp2(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Requests HTTP/2 during the TLS handshake by advertising the `h2` protocol via ALPN.

## Extension System: Pluggable Middleware, Protocol, and Compression

Wasabi allows you to extend its behaviour without modifying the core module. Three extension types can be registered per handle.

### WasabiUseProtocol

```vb
Public Sub WasabiUseProtocol(ByVal extension As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a protocol handler that receives parsed text and binary messages directly, bypassing the internal message queue. The object must implement:

```vb
Public Sub OnConnect(ByVal handle As Long)
Public Sub OnDisconnect(ByVal handle As Long)
Public Sub OnTextMessage(ByVal handle As Long, ByVal message As String)
Public Sub OnBinaryMessage(ByVal handle As Long, ByRef data() As Byte)
```

If a handler is registered, `WebSocketReceive` and `WebSocketReceiveBinary` will always return empty results for that handle.

### WasabiUseMiddleware

```vb
Public Sub WasabiUseMiddleware(ByVal extension As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a middleware object that intercepts raw byte arrays on the send and receive paths. Multiple middlewares can be registered per handle and are executed in registration order. The object must implement:

```vb
Public Sub OnConnect(ByVal handle As Long)
Public Sub OnDisconnect(ByVal handle As Long)
Public Sub OnBeforeSend(ByVal handle As Long, ByRef data() As Byte)
Public Sub OnAfterReceive(ByVal handle As Long, ByRef data() As Byte)
```

The `data()` array is passed `ByRef` and may be modified or replaced in place. Middleware sees data before compression (outbound) and after decompression (inbound).

### WasabiUseCompression

```vb
Public Sub WasabiUseCompression(ByVal extension As Object, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a compression handler that replaces the built-in `permessage-deflate` path. The object must implement:

```vb
Public Sub OnConnect(ByVal handle As Long)
Public Sub OnDisconnect(ByVal handle As Long)
Public Function Deflate(ByRef data() As Byte, ByVal windowBits As Long, ByVal contextTakeover As Boolean) As Byte()
Public Function Inflate(ByRef data() As Byte, ByVal windowBits As Long, ByVal contextTakeover As Boolean) As Byte()
```

If no handler is registered, no compression occurs. The core module is completely independent of any compression library.

### Chaining Middleware

```vb
Dim logger As New MyLogger
Dim encryptor As New MyEncryptor

WasabiUseMiddleware logger, h
WasabiUseMiddleware encryptor, h
```

Full extension specifications are available in the [extensions/](extensions/) directory.

## TCP Client

Wasabi includes a full native TCP client that shares the same connection pool and infrastructure as WebSocket. Plain TCP and TLS TCP connections are supported, with the same proxy, MTU, certificate, and timeout configuration available on WebSocket handles.

### Connection

#### TcpConnect

```vb
Public Function TcpConnect(ByVal host As String, ByVal port As Long, ByRef outHandle As Long) As Boolean
```

Opens a plain TCP connection to the specified host and port. Uses Happy Eyeballs (RFC 6555) for IPv4/IPv6 racing, automatic MTU discovery, and applies socket options on connect.

```vb
Dim h As Long
If TcpConnect("tcpbin.com", 4242, h) Then
    TcpSendText "hello" & vbCrLf, h
    TcpDisconnect h
End If
```

#### TcpConnectTLS

```vb
Public Function TcpConnectTLS(ByVal host As String, ByVal port As Long, ByRef outHandle As Long) As Boolean
```

Opens a TCP connection with TLS 1.2/1.3 via Schannel SSPI. The full TLS handshake and optional certificate validation sequence runs identically to `wss://` connections.

#### TcpDisconnect

```vb
Public Sub TcpDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Closes the TCP socket and releases all resources for the handle. Disables auto reconnect.

#### TcpIsConnected

```vb
Public Function TcpIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns `True` if the handle is active and in TCP or TCP_TLS mode.

#### TcpGetConnectionCount

```vb
Public Function TcpGetConnectionCount() As Long
```

Returns the number of currently active TCP handles (both plain and TLS).

#### TcpGetAllHandles

```vb
Public Function TcpGetAllHandles() As Long()
```

Returns an array of all currently active TCP handle indices.

### Sending Data

#### TcpSend

```vb
Public Function TcpSend(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a raw byte array. Uses `TLSSend` internally when the handle is in `MODE_TCP_TLS`, otherwise uses `send` directly.

#### TcpSendText

```vb
Public Function TcpSendText(ByVal text As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Encodes the string to UTF-8 and sends it via `TcpSend`.

#### TcpBroadcast

```vb
Public Function TcpBroadcast(ByRef data() As Byte) As Long
```

Sends a byte array to all active TCP handles. Returns the count of successful sends.

#### TcpBroadcastText

```vb
Public Function TcpBroadcastText(ByVal text As String) As Long
```

Encodes text to UTF-8 and sends it to all active TCP handles.

### Receiving Data

#### TcpReceive

```vb
Public Function TcpReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
```

Reads all currently available bytes from the receive buffer and returns them as a byte array. Returns an empty array if nothing is available. Each call drives internal maintenance (inactivity timeout, MTU probe) and reads the OS socket buffer.

#### TcpReceiveText

```vb
Public Function TcpReceiveText(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Calls `TcpReceive` and decodes the result from UTF-8.

#### TcpReceiveUntil

```vb
Public Function TcpReceiveUntil(ByVal delimiter As String, Optional ByVal timeoutMs As Long = 5000, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Blocks until the delimiter is found in the receive stream or the timeout expires. Returns all bytes up to and including the delimiter as a UTF-8 string. Bytes after the delimiter are preserved in the internal buffer for the next call.

#### TcpFlushBuffer

```vb
Public Sub TcpFlushBuffer(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Discards all bytes currently in the TCP receive buffer.

#### TcpGetPendingBytes

```vb
Public Function TcpGetPendingBytes(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the number of bytes waiting in the internal TCP receive buffer.

### Configuration

#### TcpSetNoDelay

```vb
Public Function TcpSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Controls `TCP_NODELAY`. Can be toggled at any time, even mid-connection.

#### TcpSetInactivityTimeout

```vb
Public Sub TcpSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Closes the connection if no data is received within the specified interval.

#### TcpSetReceiveTimeout

```vb
Public Sub TcpSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the timeout for internal `select()` calls.

#### TcpSetBufferSize

```vb
Public Sub TcpSetBufferSize(ByVal bufferSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Overrides the receive buffer size. Must be set before connecting. Valid range: 8192 to 16777216 bytes.

#### TcpSetPreferIPv6

```vb
Public Sub TcpSetPreferIPv6(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Biases Happy Eyeballs toward IPv6 when resolving hostnames.

#### TcpSetMTU

```vb
Public Sub TcpSetMTU(ByVal mtu As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets a static MTU value (576 to 9000). Default is 1500.

#### TcpSetAutoMTU

```vb
Public Sub TcpSetAutoMTU(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables automatic MTU discovery via `TCP_MAXSEG`.

#### TcpSetErrorDialog

```vb
Public Sub TcpSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Controls whether errors trigger a `MsgBox`.

#### TcpSetLogCallback

```vb
Public Sub TcpSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a VBA macro as a log receiver for this TCP handle.

### TLS Configuration

#### TcpSetCertValidation

```vb
Public Sub TcpSetCertValidation(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables server certificate chain validation after the TLS handshake.

#### TcpSetRevocationCheck

```vb
Public Sub TcpSetRevocationCheck(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables CRL/OCSP revocation checking.

#### TcpSetClientCert

```vb
Public Sub TcpSetClientCert(ByVal thumbprintOrSubject As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Loads a client certificate from the Windows Personal store for mTLS.

#### TcpSetClientCertPfx

```vb
Public Sub TcpSetClientCertPfx(ByVal pfxPath As String, ByVal pfxPassword As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Loads a client certificate from a PFX file for mTLS.

### Proxy Configuration

#### TcpSetProxy

```vb
Public Sub TcpSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal proxyType As Long = PROXY_TYPE_HTTP, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Configures an HTTP CONNECT or SOCKS5 proxy.

#### TcpClearProxy

```vb
Public Sub TcpClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes proxy configuration from the handle.

#### TcpAutoDiscoverProxy

```vb
Public Sub TcpAutoDiscoverProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Auto-detects the Windows system proxy via `WinHttpGetIEProxyConfigForCurrentUser`.

#### TcpGetProxyInfo

```vb
Public Function TcpGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited proxy configuration summary.

### Diagnostics and Stats

#### TcpGetStats

```vb
Public Function TcpGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited snapshot of connection metrics.

#### TcpResetStats

```vb
Public Sub TcpResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Resets all counters and updates the connected timestamp.

#### TcpGetUptime

```vb
Public Function TcpGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns seconds since the connection was established.

#### TcpGetLatency

```vb
Public Function TcpGetLatency(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the last measured RTT in milliseconds.

#### TcpGetMTUInfo

```vb
Public Function TcpGetMTUInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited summary of MTU, MSS, and optimal frame size.

#### TcpGetHost

```vb
Public Function TcpGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the connected hostname.

#### TcpGetPort

```vb
Public Function TcpGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the connected port.

#### TcpGetMode

```vb
Public Function TcpGetMode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiConnectionMode
```

Returns `MODE_TCP` or `MODE_TCP_TLS`.

#### TcpGetLastError

```vb
Public Function TcpGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
```

Returns the most recent `WasabiError` value for the handle.

#### TcpGetLastErrorCode

```vb
Public Function TcpGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the most recent native system error code.

#### TcpGetTechnicalDetails

```vb
Public Function TcpGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a human-readable description of the most recent error.

```vb
Sub TcpEchoLoop()
    Dim h As Long

    If Not TcpConnect("tcpbin.com", 4242, h) Then
        Debug.Print "Failed: " & TcpGetTechnicalDetails(h)
        Exit Sub
    End If

    TcpSetInactivityTimeout 30000, h
    TcpSetNoDelay True, h

    Dim i As Long
    For i = 1 To 5
        TcpSendText "msg_" & i & vbCrLf, h

        Dim msg As String
        Dim t As Long
        t = GetTickCount()
        Do While TickDiff(t, GetTickCount()) < 3000
            msg = TcpReceiveText(h)
            If Len(msg) > 0 Then Exit Do
            DoEvents
        Loop

        Debug.Print "Echo " & i & ": " & Trim(msg)
    Next i

    Debug.Print TcpGetStats(h)
    TcpDisconnect h
End Sub
```

## MQTT Client

Wasabi includes an MQTT client that uses the existing WebSocket transport. It supports the full MQTT 3.1.1 specification with MQTT 5 extensions, including User Properties, Reason Codes, and metadata parsing.

All MQTT functions share the same WebSocket connection handle. You must call `WebSocketConnect` with a WebSocket URL before using any MQTT function.

> [!NOTE]
> The MQTT client supports QoS 0 (at most once), QoS 1 (at least once), and QoS 2 (exactly once) for publishing, featuring an internal in-flight message queue, Packet ID generation, and full acknowledgment handshakes.

### MqttConnect

```vb
Public Function MqttConnect(ByVal clientId As String, Optional ByVal username As String, Optional ByVal password As String, Optional ByVal keepAlive As Integer = 60, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends an MQTT CONNECT packet over the established WebSocket connection.

### MqttPublish

```vb
Public Function MqttPublish(ByVal topic As String, ByVal message As String, Optional ByVal qos As Byte = 0, Optional ByVal retained As Boolean = False, Optional ByVal metaKey As String = "", Optional ByVal metaValue As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Publishes a text message to the given topic. When using QoS 1 or 2, Wasabi automatically generates a Packet ID, queues the message, and handles the complete acknowledgment handshake (`PUBACK` for QoS 1; `PUBREC`, `PUBREL`, and `PUBCOMP` for QoS 2) behind the scenes.

The optional `metaKey` and `metaValue` parameters attach MQTT 5 User Properties to the PUBLISH packet. Brokers that support MQTT 5 will forward these key/value pairs to subscribers.

### MqttSubscribe

```vb
Public Function MqttSubscribe(ByVal topic As String, Optional ByVal qos As Byte = 0, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Subscribes to a topic with the specified QoS.

### MqttUnsubscribe

```vb
Public Function MqttUnsubscribe(ByVal topic As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Removes a topic subscription.

### MqttDisconnect

```vb
Public Function MqttDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends an MQTT DISCONNECT packet and closes the MQTT session.

### MqttPingReq

```vb
Public Function MqttPingReq(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends an MQTT PINGREQ keep-alive packet.

### MqttReceive

```vb
Public Function MqttReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Polls for incoming MQTT messages. Returns a string in the format `topic|payload` when a PUBLISH packet is received, or an empty string when no message is available. For MQTT 5, the returned string may include appended User Properties (`|key=value` pairs) after the payload.

This function also returns control strings such as `[CONNACK]`, `[SUBACK]`, `[DISCONNECT]`, or `[CONNACK_ERROR]` with error details when acknowledgment packets are parsed.

## Compression and Extensions

Compression is no longer a mandatory dependency. Wasabi uses a pluggable architecture where compression is opt-in per connection. You must register a handler via `WasabiUseCompression` for any compression to occur.

The official `permessage-deflate` extension is provided as a separate class (`ExtWasabiZlib.cls`). It requires `zlib1.dll` only if you choose to use it. You can write your own compressor (LZ4, Brotli, Zstd) and register it the same way. If no compression handler is registered, `WebSocketSetDeflate` or the `DeflateEnabled` parameter will still negotiate the extension header, but no actual compression will happen.

### WebSocketSetDeflate

```vb
Public Sub WebSocketSetDeflate(ByVal enabled As Boolean, Optional ByVal contextTakeover As Boolean = True, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables `permessage-deflate` negotiation. Must be set before connecting. The actual compression algorithm is determined by the registered compression handler.

### WebSocketGetDeflateEnabled

```vb
Public Function WebSocketGetDeflateEnabled(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns `True` if `permessage-deflate` was successfully negotiated and a compression handler is active.

## Error Reference

### WasabiError Enumeration

| Code | Name | Cause |
|:---|:---|:---|
| 0 | `ERR_NONE` | No error |
| 1 | `ERR_WSA_STARTUP_FAILED` | `WSAStartup` returned a non-zero code |
| 2 | `ERR_SOCKET_CREATE_FAILED` | `socket()` returned `INVALID_SOCKET` |
| 3 | `ERR_DNS_RESOLVE_FAILED` | `gethostbyname()` returned null or WSA error 11001-11004 |
| 4 | `ERR_CONNECT_FAILED` | `connect()` failed or `select()` timed out during connection |
| 5 | `ERR_TLS_ACQUIRE_CREDS_FAILED` | `AcquireCredentialsHandle` returned a non-zero SSPI code |
| 6 | `ERR_TLS_HANDSHAKE_FAILED` | `InitializeSecurityContext` returned a fatal SSPI error |
| 7 | `ERR_TLS_HANDSHAKE_TIMEOUT` | TLS handshake loop exceeded 30 iterations or data wait timed out |
| 8 | `ERR_WEBSOCKET_HANDSHAKE_FAILED` | Could not send or receive the HTTP upgrade request |
| 9 | `ERR_WEBSOCKET_HANDSHAKE_TIMEOUT` | Server did not respond to the upgrade request within timeout |
| 10 | `ERR_SEND_FAILED` | `send()` returned zero or negative after TLS encryption |
| 11 | `ERR_RECV_FAILED` | `recv()` returned a negative value |
| 12 | `ERR_NOT_CONNECTED` | A send was attempted on a handle that is not connected |
| 13 | `ERR_ALREADY_CONNECTED` | Reserved for future use |
| 14 | `ERR_TLS_ENCRYPT_FAILED` | `EncryptMessage` returned a non-zero SSPI code |
| 15 | `ERR_TLS_DECRYPT_FAILED` | `DecryptMessage` returned a fatal SSPI code (excluding `SEC_I_RENEGOTIATE`) |
| 16 | `ERR_INVALID_URL` | URL does not begin with `ws://` or `wss://` or could not be parsed |
| 17 | `ERR_HANDSHAKE_REJECTED` | Server returned non-101 status or `Sec-WebSocket-Accept` was invalid |
| 18 | `ERR_CONNECTION_LOST` | `recv()` returned zero or an oversized frame was received |
| 19 | `ERR_INVALID_HANDLE` | Handle is out of range (reserved for future validation) |
| 20 | `ERR_MAX_CONNECTIONS` | All 64 pool slots are in use |
| 21 | `ERR_PROXY_CONNECT_FAILED` | Could not send CONNECT or proxy did not respond |
| 22 | `ERR_PROXY_AUTH_FAILED` | Proxy returned HTTP 407 (or SOCKS5 auth rejected) |
| 23 | `ERR_PROXY_TUNNEL_FAILED` | Proxy returned a non-200 status for CONNECT |
| 24 | `ERR_INACTIVITY_TIMEOUT` | No data received within the configured inactivity window |
| 25 | `ERR_CERT_LOAD_FAILED` | Failed to load client certificate from PFX or Windows store |
| 26 | `ERR_CERT_VALIDATE_FAILED` | Server certificate chain validation failed |
| 27 | `ERR_FRAGMENT_OVERFLOW` | Received fragmented message exceeds the fragment buffer size |
| 28 | `ERR_TLS_RENEGOTIATE` | Server requested TLS renegotiation (not supported) |

## Practical Patterns

### Single-Connection Workbook

```vb
Sub RunBot()
    If Not WebSocketConnect("wss://echo.websocket.org") Then
        Debug.Print "Failed to connect"
        Exit Sub
    End If

    WebSocketSetAutoReconnect True
    WebSocketSetPingInterval 25000

    Do While WebSocketIsConnected()
        Dim msg As String
        msg = WebSocketReceive()

        If msg <> "" Then
            Debug.Print "Received:", msg
            WebSocketSend "Echo: " & msg
        End If

        DoEvents
    Loop
End Sub
```

### Multi-Connection Workbook

```vb
Dim g_MarketHandle As Long
Dim g_EventHandle As Long

Sub OpenConnections()
    WebSocketAddHeader "Authorization", "Bearer token1"
    WebSocketConnect "wss://market.example.com/stream", g_MarketHandle
    WebSocketClearHeaders

    WebSocketAddHeader "Authorization", "Bearer token2"
    WebSocketConnect "wss://events.example.com/ws", g_EventHandle
    WebSocketClearHeaders

    WebSocketSetPingInterval 30000, 5000, g_MarketHandle
    WebSocketSetPingInterval 30000, 5000, g_EventHandle

    Application.OnTime Now + TimeValue("00:00:01"), "PollConnections"
End Sub

Sub PollConnections()
    Dim market As String
    Dim evt As String

    market = WebSocketReceive(g_MarketHandle)
    evt = WebSocketReceive(g_EventHandle)

    If market <> "" Then Sheet1.Range("A1").Value = market
    If evt <> "" Then Sheet1.Range("B1").Value = evt

    If WebSocketIsConnected(g_MarketHandle) Or WebSocketIsConnected(g_EventHandle) Then
        Application.OnTime Now + TimeValue("00:00:01"), "PollConnections"
    End If
End Sub
```

### Async Event-Driven Connection

```vb
' In a standard module at module level
Private g_Handler As cWasabiAsync
Private g_Handle As Long

Public Sub StartAsync()
    Set g_Handler = New cWasabiAsync
    If WebSocketConnect("wss://stream.binance.com:9443/ws/btcusdt@trade", g_Handle) Then
        WasabiUseAsync g_Handler, g_Handle
        Debug.Print "Async mode active. Callbacks fire while Excel is idle."
    End If
End Sub

Public Sub StopAsync()
    WebSocketDisconnect g_Handle
    Set g_Handler = Nothing
End Sub
```

Handler class (`cWasabiAsync`):

```vb
Option Explicit

Public Sub OnConnect(ByVal handle As Long)
    Debug.Print "Connected. Handle: " & handle
End Sub

Public Sub OnReceive(ByVal handle As Long)
    Dim msg As String
    Do While WebSocketGetPendingCount(handle) > 0
        msg = WebSocketReceive(handle)
        Debug.Print "Received: " & msg
    Loop
End Sub

Public Sub OnReadyToSend(ByVal handle As Long)
End Sub

Public Sub OnClose(ByVal handle As Long)
    Debug.Print "Connection closed. Handle: " & handle
End Sub

Public Sub OnError(ByVal handle As Long, ByVal errorCode As Long, ByVal eventType As Long)
    Debug.Print "Error on handle " & handle & " | Code: " & errorCode
End Sub
```

### MQTT IoT Dashboard with MQTT 5 User Properties

```vb
Sub StartMqttDashboard()
    Dim h As Long

    WebSocketAutoDiscoverProxy h
    If Not WebSocketConnect("wss://test.mosquitto.org:8081/mqtt", h, True, True, "mqtt") Then
        Debug.Print "Connection failed"
        Exit Sub
    End If

    WebSocketSetPingInterval 20000, 5000, h
    WebSocketSetOfflineQueueing True, h

    MqttConnect "wasabi-dashboard", h
    MqttSubscribe "sensors/temperature", 0, h

    Do
        Dim msg As String
        msg = MqttReceive(h)

        If msg <> "" And Left(msg, 1) <> "[" Then
            Dim parts() As String
            parts = Split(msg, "|", 3)
            If parts(0) = "sensors/temperature" Then
                Sheet1.Cells(2, 1).Value = Now()
                Sheet1.Cells(2, 2).Value = parts(1)
                ' parts(2) may contain "key=value" from MQTT 5 User Properties
            End If
        End If
        DoEvents
    Loop While WebSocketIsConnected(h)

    MqttDisconnect h
    WebSocketDisconnect h
End Sub
```

## Operational Caveats

> [!WARNING]
> Wasabi operates entirely on the VBA main thread. There are no background threads. All socket activity, maintenance, and reconnect logic runs when your code explicitly calls a Wasabi function or when the Windows message pump dispatches an async event.

> [!WARNING]
> In polling mode, automatic heartbeat, ping scheduling, inactivity timeout detection, and auto reconnect triggering are all driven by polling calls. If your code stops calling receive functions, these features stop functioning.

> [!WARNING]
> In async mode, the handler object must remain alive for the entire duration of the connection. Store it at module or workbook level. Always call `WebSocketDisconnect` or `WebSocketDisconnectAll` before resetting the VBA project or closing the workbook.

> [!WARNING]
> Queue capacity is fixed at 512 messages per type per connection. Under sustained high message rates, messages will be dropped without error if the queue is not drained fast enough.

> [!CAUTION]
> Custom headers, subprotocol, proxy configuration, buffer sizes, MTU settings, security options, and extension registrations must be set before calling `WebSocketConnect`. Changes after connection have no effect on the active session.

> [!CAUTION]
> `WebSocketDisconnect` always disables auto reconnect. This is intentional and cannot be overridden.

> [!NOTE]
> The pipe-delimited format returned by diagnostic functions (`GetStats`, `GetReconnectInfo`, `GetProxyInfo`, `GetMTUInfo`) is intended for human-readable diagnostics. Do not build parsing logic that depends on field order or format stability across future versions.

> [!NOTE]
> TCP handles do not support WebSocket-specific features such as message queuing, ping scheduling, MQTT, `permessage-deflate`, offline queueing, or protocol handlers. These features are exclusive to `MODE_WEBSOCKET` handles.

> [!NOTE]
> `TcpReceive` and `TcpReceiveText` return all bytes currently available in the buffer in a single call. TCP is a stream protocol with no message boundaries. Use `TcpReceiveUntil` when your protocol uses delimiters, or implement framing logic in your application layer.
