# API Reference

This document describes the complete public interface of `Wasabi.bas`.

Wasabi is a native WebSocket and WSS client for VBA built on raw Winsock and Schannel. It implements the WebSocket protocol as defined by RFC 6455 and handles TLS through the Windows SSPI Schannel provider. The public API is minimal at the call site, but understanding the runtime model is essential for reliable usage.

## Core Concepts

### Handles and the Connection Pool

Wasabi maintains an internal pool of up to 64 simultaneous connections. Each connection is identified by an integer handle, which is an index into this pool. Handles are allocated automatically when you call `WebSocketConnect` and are released when you call `WebSocketDisconnect`.

Most public functions accept an optional `handle` parameter. The parameter type is `Long`. Valid handles range from `0` to `63` inclusive.

> [!IMPORTANT]
> Do not assume that handle `0` is invalid or uninitialized. `0` is a fully valid connection handle.

The recommended pattern for multi connection applications is to always declare, store, and pass handles explicitly.

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

VBA is a single-threaded language. There is one execution thread shared between your code, the Office runtime, and all Wasabi operations. Background threads are not available in the VBA runtime.

This means Wasabi cannot push messages to your code automatically when they arrive. Instead, received data accumulates in internal buffers and is processed when you call one of the receive functions.

Each call to `WebSocketReceive` or `WebSocketReceiveAll` triggers the following internal sequence:

1. Calls internal maintenance to check ping intervals and inactivity timeouts
2. Processes any already-buffered decoded frames
3. Checks the OS socket buffer for available data using `ioctlsocket` with `FIONREAD`
4. If data is available, reads it with `recv`
5. If TLS is active, passes the raw bytes through `DecryptMessage`
6. Parses WebSocket frames and routes each by opcode
7. Enqueues complete text and binary messages into their respective ring buffers
8. Returns the oldest queued message

> [!IMPORTANT]
> Your code must call a receive function regularly for any of the following to work: message delivery, automatic ping scheduling, inactivity timeout detection, and auto reconnect triggering.

### Message Queues and Capacity

Wasabi uses two separate circular queues per connection, one for text messages and one for binary messages. Each has a fixed capacity of 512 entries.

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

## Connection Management

### WebSocketConnect

```vb
Public Function WebSocketConnect(ByVal url As String, Optional ByRef outHandle As Long = -1) As Boolean
```

Opens a new WebSocket connection to the specified URL.

This function executes the complete connection sequence in order:

1. Initializes Winsock if not already active
2. Allocates a slot in the connection pool
3. Parses the URL into host, port, path, and scheme
4. Resolves the hostname via `gethostbyname` or `inet_addr` if a literal IP is supplied
5. Creates a non-blocking TCP socket
6. Initiates connection with `connect` and waits for completion using `select`
7. Applies `TCP_NODELAY`, `SO_KEEPALIVE`, `SO_RCVBUF`, and `SO_SNDBUF` socket options
8. If proxy is configured, sends an HTTP CONNECT request (or SOCKS5) and verifies the tunnel
9. If the scheme is `wss://`, performs a TLS handshake via Schannel SSPI
10. If configured, validates the server certificate chain
11. Queries `SecPkgContext_StreamSizes` to determine TLS record framing parameters
12. Sends the HTTP/1.1 WebSocket upgrade request
13. Validates the `Sec-WebSocket-Accept` header value using SHA-1 and Base64
14. Marks the connection as active and records the connection timestamp

#### Parameters

* `url`: a WebSocket URL beginning with `ws://` or `wss://`. Custom ports may be specified inline, for example `wss://api.example.com:8443/ws`.
* `outHandle`: receives the allocated integer handle. Set to `-1` on failure.

#### Returns

`True` if the connection was fully established. `False` if any step in the sequence failed.

#### Example

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
> Configure all options (proxy, custom headers, subprotocol, buffer sizes, MTU, no-delay) before calling `WebSocketConnect`. These settings are applied during the connection sequence and cannot be changed mid-connection.

> [!CAUTION]
> The connection pool has a limit of 64 simultaneous connections. Attempting to connect beyond this limit returns `ERR_MAX_CONNECTIONS`.

### WebSocketDisconnect

```vb
Public Sub WebSocketDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Gracefully terminates a connection.

This function disables auto reconnect for the handle, sends a WebSocket Close frame with status code 1000 if the connection is active, closes the underlying TCP socket, releases the Schannel security context and credential handle, and clears all internal connection state.

If no other connections remain active after this call, Wasabi also calls `WSACleanup` to release Winsock resources.

#### Example

```vb
WebSocketDisconnect h
```

```vb
' Disconnect and check all connections are gone
WebSocketDisconnect h

If WebSocketGetConnectionCount() = 0 Then
    Debug.Print "All connections closed"
End If
```

> [!NOTE]
> Calling `WebSocketDisconnect` always disables auto reconnect for the target handle. This is intentional. Disconnecting means you chose to close the connection, so auto reconnect would be unexpected behavior.

### WebSocketDisconnectAll

```vb
Public Sub WebSocketDisconnectAll()
```

Iterates through the entire pool and disconnects every active connection.

#### Example

```vb
' Cleanup on workbook close
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    WebSocketDisconnectAll
End Sub
```

### WebSocketIsConnected

```vb
Public Function WebSocketIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns the current connected state of the handle.

This is a lightweight state check. It reads the internal `Connected` flag without touching the socket.

> [!NOTE]
> This function reports the last known state. It does not actively probe the socket. A connection may have dropped between the last receive call and this check.

#### Example

```vb
Do While WebSocketIsConnected(h)
    Dim msg As String
    msg = WebSocketReceive(h)
    If msg <> "" Then Debug.Print msg
    DoEvents
Loop

Debug.Print "Connection ended"
```

### WebSocketSendClose

```vb
Public Function WebSocketSendClose(Optional ByVal code As Integer = 1000, Optional ByVal reason As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a WebSocket Close control frame and marks the connection as disconnected.

The status code occupies the first two bytes of the Close frame payload in big-endian byte order, followed by the UTF-8 encoded reason string.

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

#### Example

```vb
WebSocketSendClose 1001, "Client shutting down", h
```

> [!NOTE]
> `WebSocketDisconnect` calls `WebSocketSendClose` internally. Call `WebSocketSendClose` directly only if you need to specify a custom close code or reason before cleanup.

## Sending Data

### WebSocketSend

```vb
Public Function WebSocketSend(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a text message.

Internally, Wasabi converts the VBA Unicode string to UTF-8 using `WideCharToMultiByte`, then constructs a masked WebSocket text frame. The payload length is encoded using the appropriate tier:

* 7-bit for payloads up to 125 bytes
* 16-bit with prefix `0x7E` for payloads up to 65535 bytes
* 64-bit with prefix `0x7F` for larger payloads

The frame is transmitted via `send` in a loop to handle partial writes.

#### Returns

`True` if all frame bytes were successfully written to the socket.

#### Example

```vb
If Not WebSocketSend("Hello from VBA", h) Then
    Debug.Print "Send failed:", WebSocketGetTechnicalDetails(h)
End If
```

```vb
' Sending a JSON payload
Dim payload As String
payload = "{""event"":""subscribe"",""channel"":""prices""}"

If WebSocketSend(payload, h) Then
    Debug.Print "Subscribed"
End If
```

> [!IMPORTANT]
> Frames are masked client-side using a random 4-byte key generated by `CryptGenRandom` as required by RFC 6455. This is done automatically and transparently.

> [!CAUTION]
> Sending on a disconnected handle returns `False` and logs `ERR_NOT_CONNECTED`. The connection state is not changed.

### WebSocketSendBinary

```vb
Public Function WebSocketSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a binary message using WebSocket opcode `0x02`.

The byte array is wrapped in a masked binary frame using the same length encoding tiers as text messages.

#### Example

```vb
' Sending a raw byte sequence
Dim payload(0 To 7) As Byte
payload(0) = &H01
payload(1) = &H02
payload(2) = &HFF
payload(3) = &H00
payload(4) = &HAB
payload(5) = &HCD
payload(6) = &HEF
payload(7) = &H42

If WebSocketSendBinary(payload, h) Then
    Debug.Print "Binary sent"
End If
```

```vb
' Reading a file and sending it as binary
Dim fileNum As Integer
Dim fileSize As Long
Dim fileData() As Byte

fileNum = FreeFile
Open "C:\data\packet.bin" For Binary As #fileNum
fileSize = LOF(fileNum)
ReDim fileData(0 To fileSize - 1)
Get #fileNum, , fileData
Close #fileNum

WebSocketSendBinary fileData, h
```

> [!CAUTION]
> Empty arrays return `True` immediately and do not transmit any frame. Ensure the array is properly allocated before passing it.

### WebSocketBroadcast

```vb
Public Function WebSocketBroadcast(ByVal message As String) As Long
```

Sends the same text message to every active connection in the pool.

Returns the count of connections that received the message successfully.

#### Example

```vb
Dim delivered As Long
delivered = WebSocketBroadcast("{""type"":""notification"",""text"":""Server restarting""}")
Debug.Print "Delivered to", delivered, "of", WebSocketGetConnectionCount(), "connections"
```

### WebSocketBroadcastBinary

```vb
Public Function WebSocketBroadcastBinary(ByRef data() As Byte) As Long
```

Sends the same binary payload to every active connection.

#### Example

```vb
Dim signal(0 To 1) As Byte
signal(0) = &HFF
signal(1) = &H00

Dim delivered As Long
delivered = WebSocketBroadcastBinary(signal)
Debug.Print "Binary broadcast delivered to", delivered, "connections"
```

### Advanced Sending

#### WebSocketSendBatch

```vb
Public Function WebSocketSendBatch(ByRef messages() As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends multiple text messages in a single TCP write (or a minimal number of writes) to reduce system call overhead. All messages are packed into a contiguous byte buffer before transmission.

Returns `True` if the entire batch was sent successfully.

#### WebSocketSendBatchBinary

```vb
Public Function WebSocketSendBatchBinary(ByRef messages() As Variant, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends multiple binary payloads in a single TCP write (or a minimal number of writes). Each element of the `messages` array must be a `Byte()` array.

Returns `True` if the entire batch was sent successfully.

#### WebSocketSendMTUAware

```vb
Public Function WebSocketSendMTUAware(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a text message using WebSocket fragmentation automatically tuned to the connection's current MTU. If the message fits in one frame, it behaves identically to `WebSocketSend`. Otherwise, the message is split into fragments sized to the `OptimalFrameSize` and sent as a sequence of continuation frames.

Enabling `AutoMTU` (default) allows this function to use the dynamic MTU calculated from the TCP MSS; otherwise it uses the static MTU set via `WebSocketSetMTU`.

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

Each invocation drives the full internal polling and maintenance cycle described in the Core Concepts section.

#### Example

```vb
' Fire and forget receive
Dim msg As String
msg = WebSocketReceive(h)
If msg <> "" Then
    Debug.Print "Got:", msg
End If
```

```vb
' Blocking wait for one response with timeout
Dim msg As String
Dim startTick As Long

startTick = GetTickCount()

Do
    msg = WebSocketReceive(h)
    If msg <> "" Then Exit Do
    If GetTickCount() - startTick > 5000 Then
        Debug.Print "Timeout waiting for response"
        Exit Do
    End If
    DoEvents
Loop

Debug.Print msg
```

```vb
' Continuous listening loop with Application.OnTime
Public Sub Tick()
    Dim msg As String
    msg = WebSocketReceive(m_Handle)

    If msg <> "" Then
        Sheet1.Cells(m_Row, 1).Value = Now()
        Sheet1.Cells(m_Row, 2).Value = msg
        m_Row = m_Row + 1
    End If

    If WebSocketIsConnected(m_Handle) Then
        Application.OnTime Now + TimeValue("00:00:01"), "Tick"
    End If
End Sub
```

> [!WARNING]
> An empty return value is ambiguous. It can mean the queue is empty or that the server sent a zero-length text frame. If zero-length frames are meaningful in your protocol, use `WebSocketGetPendingCount` to distinguish the two cases before calling `WebSocketReceive`.

> [!TIP]
> Use `Application.OnTime` for dashboard scenarios where you want the spreadsheet to remain fully interactive between message checks.

### WebSocketReceiveAll

```vb
Public Function WebSocketReceiveAll(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String()
```

Drains the complete text queue and returns all pending messages as a string array.

This is more efficient than calling `WebSocketReceive` in a loop when you expect bursts of messages.

#### Example

```vb
Dim messages() As String
Dim i As Long

messages = WebSocketReceiveAll(h)

If UBound(messages) >= LBound(messages) Then
    For i = LBound(messages) To UBound(messages)
        Debug.Print i, messages(i)
    Next i
Else
    Debug.Print "No messages"
End If
```

> [!CAUTION]
> VBA array bounds can behave unexpectedly when the array is empty. Always guard iteration with an `UBound >= LBound` check.

### WebSocketReceiveBinary

```vb
Public Function WebSocketReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
```

Returns the oldest binary message from the queue.

#### Example

```vb
Dim data() As Byte
Dim i As Long

data = WebSocketReceiveBinary(h)

If Not Not data Then
    Debug.Print "Received", UBound(data) - LBound(data) + 1, "bytes"
    For i = LBound(data) To UBound(data)
        Debug.Print "Byte", i, "=", Hex(data(i))
    Next i
End If
```

> [!CAUTION]
> In VBA, an empty byte array can cause errors if you access it without checking. The `Not Not data` idiom evaluates to `False` when the array has not been initialized or is empty.

### WebSocketReceiveBinaryCheck

```vb
Public Function WebSocketReceiveBinaryCheck(ByRef outData() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Populates `outData` with the oldest binary message and returns `True` if data was available.

#### Example

```vb
Dim data() As Byte

If WebSocketReceiveBinaryCheck(data, h) Then
    Debug.Print "Received", UBound(data) - LBound(data) + 1, "bytes"
Else
    Debug.Print "No binary message"
End If
```

> [!TIP]
> Prefer `WebSocketReceiveBinaryCheck` over `WebSocketReceiveBinary` when you want explicit control flow based on whether data was available.

### Zero-Copy Receiving

Zero-copy functions allow direct access to the internal data buffers without creating copies. They must be enabled per-connection via `WebSocketSetZeroCopy`.

#### WebSocketSetZeroCopy

```vb
Public Sub WebSocketSetZeroCopy(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables zero-copy mode for text and binary reception. When active, `WebSocketReceiveZeroCopy` and `WebSocketReceiveBinaryZeroCopy` expose internal pointers instead of returning new strings or arrays.

#### WebSocketReceiveZeroCopy

```vb
Public Function WebSocketReceiveZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns a pointer to the internal string buffer of the next text message, along with its length in characters. The pointer is valid until the next call that modifies the queue. Returns `True` if a message was available.

#### WebSocketReceiveBinaryZeroCopy

```vb
Public Function WebSocketReceiveBinaryZeroCopy(ByRef outPtr As Long, ByRef outLen As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns a pointer to the internal byte array of the next binary message, along with its length in bytes. The pointer is valid until the next call that modifies the binary queue.

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

#### Example

```vb
' Monitor queue health
Sub CheckQueueHealth(ByVal h As Long)
    Dim textPending As Long
    Dim textCapacity As Long
    Dim binPending As Long
    Dim binCapacity As Long

    textPending = WebSocketGetPendingCount(h)
    textCapacity = WebSocketGetQueueCapacity(h)
    binPending = WebSocketGetBinaryPendingCount(h)
    binCapacity = WebSocketGetBinaryQueueCapacity(h)

    Debug.Print "Text queue:", textPending, "pending,", textCapacity, "remaining"
    Debug.Print "Binary queue:", binPending, "pending,", binCapacity, "remaining"

    If textCapacity < 50 Then
        Debug.Print "Warning: text queue is nearly full"
    End If
End Sub
```

### WebSocketPeek

```vb
Public Function WebSocketPeek(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the next queued text message without removing it from the queue.

#### Example

```vb
' Inspect the next message before processing
Dim preview As String
preview = WebSocketPeek(h)

If InStr(preview, """error""") > 0 Then
    Debug.Print "Next message is an error:", preview
Else
    Debug.Print "Next message looks normal"
End If

' Now actually consume it
Dim msg As String
msg = WebSocketReceive(h)
```

### WebSocketFlushQueue

```vb
Public Sub WebSocketFlushQueue(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Immediately clears all queued text and binary messages for the specified connection.

#### Example

```vb
' Discard stale messages after reconnect
WebSocketFlushQueue h
Debug.Print "Queue cleared"
```

> [!WARNING]
> This operation is irreversible. All queued messages are discarded immediately. Data still pending in the OS socket buffer is not affected and may appear in subsequent receive calls.

## Control Frames and Heartbeat

### WebSocketSendPing

```vb
Public Function WebSocketSendPing(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a Ping control frame with an optional text payload up to 125 bytes.

The Ping frame uses opcode `0x09`. The server is expected to respond with a Pong containing the same payload.

#### Example

```vb
' Manual ping with timestamp payload
Dim ts As String
ts = CStr(Now())

If WebSocketSendPing(ts, h) Then
    Debug.Print "Ping sent at", ts
End If
```

### WebSocketSendPong

```vb
Public Function WebSocketSendPong(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a Pong control frame with an optional payload.

> [!NOTE]
> Wasabi automatically responds to incoming server Pings with a matching Pong during frame processing. You only need to call this manually for unsolicited pongs or for diagnostic purposes.

### WebSocketSetPingInterval

```vb
Public Sub WebSocketSetPingInterval(ByVal intervalMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the interval in milliseconds for automatic Ping frames. Set to `0` to disable.

#### Example

```vb
' Send a ping every 30 seconds to keep the connection alive
WebSocketSetPingInterval 30000, h
```

```vb
' Disable automatic pings
WebSocketSetPingInterval 0, h
```

> [!IMPORTANT]
> Automatic pings are evaluated and dispatched during the maintenance phase that runs inside each `WebSocketReceive` call. If your code stops calling receive functions, automatic pings also stop. This is not a background timer.

### WebSocketGetLatency

```vb
Public Function WebSocketGetLatency(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the most recent round-trip time (RTT) in milliseconds, measured from the last Ping frame sent (manual or automatic) to the corresponding Pong frame received.

Returns `0` if no RTT measurement has been recorded yet.

#### Example

```vb
' Send a ping and retrieve the latency
WebSocketSendPing "", h
Dim latency As Long
latency = WebSocketGetLatency(h)
Debug.Print "Current latency: " & latency & "ms"
```

## Reconnect and Reliability

### WebSocketSetAutoReconnect

```vb
Public Sub WebSocketSetAutoReconnect(ByVal enabled As Boolean, Optional ByVal maxAttempts As Long = DEFAULT_RECONNECT_MAX_ATTEMPTS, Optional ByVal baseDelayMs As Long = DEFAULT_RECONNECT_BASE_DELAY_MS, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables automatic reconnection using exponential backoff.

When a session loss is detected during polling, Wasabi saves all connection settings, cleans up resources, waits for the calculated delay, and invokes `ConnectHandle` with the original URL to re-establish the session.

The delay between attempts doubles on each failure:

* Attempt 1: `baseDelayMs`
* Attempt 2: `baseDelayMs * 2`
* Attempt 3: `baseDelayMs * 4`
* and so on, capped at 30 seconds

The following settings are saved and restored across reconnect:

* original URL
* proxy host, port, user, password, type
* custom HTTP headers
* subprotocol
* NoDelay flag
* log callback name
* error dialog setting
* ping interval
* receive timeout
* inactivity timeout
* buffer sizes

#### Example

```vb
' Enable reconnect with up to 10 attempts, starting at 2 second delay
WebSocketSetAutoReconnect True, 10, 2000, h
```

```vb
' Disable auto reconnect
WebSocketSetAutoReconnect False, , , h
```

> [!WARNING]
> Reconnect attempts run on the main VBA thread. The delay period uses `DoEvents` internally, which yields to the Windows message pump but does not fully release the thread. UI interactions during this window may behave unexpectedly.

> [!CAUTION]
> Calling `WebSocketDisconnect` disables auto reconnect before closing. This is intentional. A manual disconnect should not trigger reconnection.

### WebSocketGetReconnectInfo

```vb
Public Function WebSocketGetReconnectInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the current reconnect configuration and attempt count as a pipe-delimited string.

#### Format

```text
AutoReconnect=1|Attempts=2|MaxAttempts=10|BaseDelayMs=2000
```

#### Example

```vb
Debug.Print WebSocketGetReconnectInfo(h)
```

## Buffer and Performance Configuration

### WebSocketSetBufferSizes

```vb
Public Sub WebSocketSetBufferSizes(ByVal bufferSize As Long, ByVal fragmentSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Overrides the default sizes for the receive buffer and the fragment reassembly buffer.

The receive buffer stores raw incoming bytes before TLS decryption or WebSocket frame parsing. The fragment buffer stores incomplete fragmented message data across multiple continuation frames.

Default size for both is 256KB (262144 bytes).

#### Valid ranges

* Receive buffer: 8192 to 16777216 bytes (8KB to 16MB)
* Fragment buffer: 4096 to 16777216 bytes (4KB to 16MB)

#### Example

```vb
' Increase buffers for a high-throughput stream
WebSocketSetBufferSizes 1048576, 1048576, h
WebSocketConnect "wss://stream.example.com/data", h
```

```vb
' Reduce buffers for low-memory environments
WebSocketSetBufferSizes 32768, 16384, h
WebSocketConnect "wss://lightweight.example.com/ws", h
```

> [!IMPORTANT]
> `WebSocketSetBufferSizes` must be called before `WebSocketConnect`. Buffer allocation happens at connection time.

> [!WARNING]
> Values outside the valid range are silently ignored. No error is raised and the default size is retained.

### WebSocketSetNoDelay

```vb
Public Function WebSocketSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Controls the `TCP_NODELAY` socket option, which disables Nagle's algorithm when enabled.

When Nagle's algorithm is active, the TCP stack coalesces small outgoing packets to improve throughput. When `TCP_NODELAY` is set, packets are sent immediately regardless of size.

#### Example

```vb
' Enable low-latency mode before connecting
WebSocketSetNoDelay True, h
WebSocketConnect "wss://realtime.example.com/ws", h
```

```vb
' Apply after connecting
WebSocketSetNoDelay True, h
Debug.Print "NoDelay applied:", WebSocketSetNoDelay(True, h)
```

> [!TIP]
> Enable `TCP_NODELAY` when your protocol sends many small messages and latency is more important than bandwidth efficiency. This is common in trading systems, real-time telemetry, and chat applications.

## MTU Management

Wasabi can automatically detect the path MTU and adjust WebSocket frame fragmentation to avoid IP fragmentation. This is controlled via the AutoMTU flag and a configurable MTU value.

### WebSocketSetMTU

```vb
Public Sub WebSocketSetMTU(ByVal mtu As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the static MTU value for frame sizing calculations. Valid range: 576 to 9000. The default is 1500. Must be set before `WebSocketConnect`.

### WebSocketGetMTU

```vb
Public Function WebSocketGetMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the current MTU value used for frame sizing.

### WebSocketSetAutoMTU

```vb
Public Sub WebSocketSetAutoMTU(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables automatic MTU discovery. When enabled (default), Wasabi probes the TCP MSS on each new connection and periodically thereafter. The resulting MTU is used to fragment outgoing frames optimally.

### WebSocketGetOptimalFrameSize

```vb
Public Function WebSocketGetOptimalFrameSize(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the calculated optimal frame payload size (in bytes) for the current MTU and TLS overhead configuration. This is the maximum payload that can be sent in a single WebSocket frame without causing IP fragmentation.

### WebSocketGetMTUInfo

```vb
Public Function WebSocketGetMTUInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited summary of the current MTU state:

```text
MTU=1500|MSS=1460|OptimalFrame=1374|AutoMTU=Yes|ProbeEnabled=Yes
```

### WebSocketProbeMTU

```vb
Public Sub WebSocketProbeMTU(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Forces an immediate MTU probe on the active socket. If the detected TCP MSS differs from the current MTU, the optimal frame size is recalculated. This is usually called automatically by the maintenance cycle when AutoMTU is active.

## Proxy Configuration

### WebSocketSetProxy

```vb
Public Sub WebSocketSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal proxyType As Long = PROXY_TYPE_HTTP, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Configures an HTTP CONNECT or SOCKS5 proxy for the connection.

`proxyType` can be `PROXY_TYPE_HTTP` (0) or `PROXY_TYPE_SOCKS5` (1). When a proxy is configured, Wasabi connects to the proxy host instead of the WebSocket server, establishes the tunnel, and then performs TLS and WebSocket handshaking through that tunnel.

Authentication for HTTP proxies uses the HTTP Basic scheme. SOCKS5 supports username/password authentication if credentials are provided.

#### Example without authentication

```vb
WebSocketSetProxy "proxy.company.local", 8080, , , , h
WebSocketConnect "wss://api.example.com/ws", h
```

#### Example with authentication

```vb
WebSocketSetProxy "proxy.company.local", 8080, "domain\user", "p@ssw0rd", PROXY_TYPE_HTTP, h
WebSocketConnect "wss://api.example.com/ws", h
```

> [!IMPORTANT]
> Proxy settings must be configured before `WebSocketConnect`.

> [!WARNING]
> HTTP proxy authentication uses Basic, which transmits Base64-encoded credentials. Use only over TLS-secured tunnels.

> [!CAUTION]
> If the proxy returns HTTP 407, Wasabi reports `ERR_PROXY_AUTH_FAILED`. If the CONNECT request returns any status other than 200, Wasabi reports `ERR_PROXY_TUNNEL_FAILED`. SOCKS5 failures map to `ERR_PROXY_CONNECT_FAILED` or `ERR_PROXY_AUTH_FAILED`.

### WebSocketSetProxyNtlm

```vb
Public Sub WebSocketSetProxyNtlm(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables NTLM/Kerberos authentication for HTTP proxies. When active and the proxy returns `407 Proxy Authentication Required` with `Proxy-Authenticate: NTLM` (or `Negotiate`), Wasabi performs a full SSPI NTLM handshake to authenticate with the current Windows credentials.

This function only affects HTTP proxies; SOCKS5 authentication continues to use the username and password supplied in `WebSocketSetProxy`.

> [!NOTE]
> This feature uses the SSPI `"NTLM"` package and the credentials of the
> currently logged‑on Windows user. No explicit username or password is
> required for the NTLM handshake.

### WebSocketClearProxy

```vb
Public Sub WebSocketClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes all proxy settings for the specified connection handle.

#### Example

```vb
WebSocketClearProxy h
```

### WebSocketGetProxyInfo

```vb
Public Function WebSocketGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a summary of the current proxy configuration.

#### Format when proxy is enabled

```text
Type=HTTP|Host=proxy.company.local|Port=8080|Auth=Yes
```

#### Format when proxy is disabled

```text
Disabled
```

#### Example

```vb
Debug.Print WebSocketGetProxyInfo(h)
```

## Handshake Customization

### WebSocketAddHeader

```vb
Public Sub WebSocketAddHeader(ByVal headerName As String, ByVal headerValue As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Adds a custom HTTP header to the WebSocket upgrade request.

Custom headers are appended to the handshake after the standard WebSocket headers and before the final blank line.

#### Example

```vb
' Bearer token authentication
WebSocketAddHeader "Authorization", "Bearer eyJhbGciOiJIUzI1NiJ9...", h

' Custom client identification
WebSocketAddHeader "X-Client-Id", "ExcelDashboard-v1", h
WebSocketAddHeader "X-Session-Id", "abc123", h

WebSocketConnect "wss://api.example.com/ws", h
```

> [!IMPORTANT]
> Headers must be added before `WebSocketConnect`. They are sent as part of the HTTP upgrade request.

> [!NOTE]
> Standard WebSocket headers such as `Host`, `Upgrade`, `Connection`, `Sec-WebSocket-Key`, and `Sec-WebSocket-Version` are always included automatically. Do not duplicate them via `WebSocketAddHeader`.

### WebSocketClearHeaders

```vb
Public Sub WebSocketClearHeaders(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes all custom headers added via `WebSocketAddHeader`.

#### Example

```vb
' Reset before a reconnect with different credentials
WebSocketClearHeaders h
WebSocketAddHeader "Authorization", "Bearer newtoken", h
```

### WebSocketSetSubProtocol

```vb
Public Sub WebSocketSetSubProtocol(ByVal protocol As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the value sent in the `Sec-WebSocket-Protocol` header during the handshake.

#### Example

```vb
' GraphQL over WebSocket
WebSocketSetSubProtocol "graphql-transport-ws", h
WebSocketConnect "wss://api.example.com/graphql", h
```

```vb
' STOMP messaging protocol
WebSocketSetSubProtocol "v12.stomp", h
WebSocketConnect "wss://broker.example.com/stomp", h
```

> [!IMPORTANT]
> The subprotocol must be set before `WebSocketConnect`.

### WebSocketGetSubProtocol

```vb
Public Function WebSocketGetSubProtocol(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the configured subprotocol string.

#### Example

```vb
Debug.Print "Subprotocol:", WebSocketGetSubProtocol(h)
```

## Timeout Configuration

### WebSocketSetReceiveTimeout

```vb
Public Sub WebSocketSetReceiveTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the timeout in milliseconds for internal `select()` calls used when waiting for data during handshake and TLS operations.

Default: 5000ms.

#### Example

```vb
' Allow 15 seconds for slow servers
WebSocketSetReceiveTimeout 15000, h
```

> [!NOTE]
> This timeout applies to blocking wait operations during the connection sequence, not to the non-blocking polling that happens in `WebSocketReceive`.

### WebSocketSetInactivityTimeout

```vb
Public Sub WebSocketSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the maximum allowed period without receiving any data from the server. If this duration elapses without any incoming bytes, the connection is treated as stale and closed.

If auto reconnect is enabled, a reconnect attempt begins immediately after the timeout closes the connection.

#### Example

```vb
' Close if server goes silent for more than 60 seconds
WebSocketSetInactivityTimeout 60000, h
```

```vb
' Combine with auto reconnect for resilient monitoring
WebSocketSetAutoReconnect True, 5, 1000, h
WebSocketSetInactivityTimeout 45000, h
```

> [!IMPORTANT]
> Inactivity is measured from the last time data was received at the socket level, not from the last time your code called `WebSocketReceive`. The timestamp is updated internally whenever bytes arrive.

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

Returns the most recent native system error code. This is typically a WSA error code (for Winsock failures) or an SSPI status code (for TLS failures), displayed in hexadecimal.

### WebSocketGetTechnicalDetails

```vb
Public Function WebSocketGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a detailed technical description of the most recent error, including function name, error code, and context.

### WebSocketGetErrorDescription

```vb
Public Function WebSocketGetErrorDescription(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a single human-readable string combining the error category, optional system code, and technical details. This is the recommended function for user-facing error logging.

#### Example

```vb
If Not WebSocketConnect("wss://bad.example.com/ws", h) Then
    Debug.Print WebSocketGetErrorDescription(h)
End If
```

#### Example output

```text
TLS handshake failed [0x80090326] - InitializeSecurityContext failed with SSPI error 0x80090326
```

### WebSocketGetStats

```vb
Public Function WebSocketGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a snapshot of connection metrics as a pipe-delimited string.

#### Format

```text
BytesSent=x|BytesReceived=x|MessagesSent=x|MessagesReceived=x|UptimeSeconds=x|Queued=x|BinaryQueued=x|NoDelay=x|Proxy=x
```

#### Example

```vb
Debug.Print WebSocketGetStats(h)
```

#### Example output

```text
BytesSent=4096|BytesReceived=12288|MessagesSent=8|MessagesReceived=24|UptimeSeconds=37|Queued=0|BinaryQueued=2|NoDelay=1|Proxy=none
```

> [!NOTE]
> `BytesSent` and `BytesReceived` count raw bytes at the socket level, including WebSocket frame overhead and TLS record framing. They do not represent application payload sizes alone.

### WebSocketResetStats

```vb
Public Sub WebSocketResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Resets all counters to zero and updates the connected timestamp to the current tick.

#### Example

```vb
' Begin a measurement window
WebSocketResetStats h

' ... perform some operations ...

Debug.Print WebSocketGetStats(h)
```

### WebSocketGetUptime

```vb
Public Function WebSocketGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns how many seconds the connection has been active since it was established or since stats were last reset.

#### Example

```vb
Debug.Print "Uptime:", WebSocketGetUptime(h), "seconds"
```

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

#### Example

```vb
Debug.Print WebSocketGetHost(h), WebSocketGetPort(h), WebSocketGetPath(h)
```

### WebSocketGetConnectionCount

```vb
Public Function WebSocketGetConnectionCount() As Long
```

Returns the total number of currently active connections.

#### Example

```vb
Debug.Print "Active connections:", WebSocketGetConnectionCount()
```

### WebSocketGetAllHandles

```vb
Public Function WebSocketGetAllHandles() As Long()
```

Returns an array of all currently active handles.

#### Example

```vb
Dim handles() As Long
Dim i As Long

handles = WebSocketGetAllHandles()

If UBound(handles) >= LBound(handles) Then
    For i = LBound(handles) To UBound(handles)
        Debug.Print "Handle", handles(i), "uptime:", WebSocketGetUptime(handles(i)), "s"
    Next i
End If
```

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

Returns a pipe-delimited summary of the last close event:

```text
Code=1000|Description=Normal Closure|Reason=|InitiatedByUs=Yes
```

## Logging and User Feedback

### WebSocketSetLogCallback

```vb
Public Sub WebSocketSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a VBA macro as a log receiver. Wasabi calls this macro using `Application.Run` whenever it emits an internal diagnostic message.

The macro must be accessible at module level and must accept exactly one `String` argument.

#### Example

```vb
' Define the callback in a standard module
Public Sub WasabiLogger(ByVal msg As String)
    Dim logRow As Long
    logRow = Sheet2.Cells(Rows.Count, 1).End(xlUp).Row + 1
    Sheet2.Cells(logRow, 1).Value = Now()
    Sheet2.Cells(logRow, 2).Value = msg
End Sub

' Register it
Sub StartConnection()
    Dim h As Long
    WebSocketConnect "wss://echo.websocket.org", h
    WebSocketSetLogCallback "WasabiLogger", h
End Sub
```

> [!IMPORTANT]
> The callback is invoked using `Application.Run`, which requires the macro to be in a standard module and be callable by name as a string. Class module methods and private procedures are not accessible this way.

> [!WARNING]
> Heavy computation or UI operations inside the callback will affect application responsiveness, since the callback runs synchronously on the main thread.

### WebSocketSetErrorDialog

```vb
Public Sub WebSocketSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Controls whether connection errors trigger a `MsgBox` dialog. Disabled by default.

#### Example

```vb
' Enable for interactive use
WebSocketSetErrorDialog True, h
```

> [!TIP]
> In unattended or production workbooks, leave dialogs disabled and use `WebSocketGetLastError` and `WebSocketGetTechnicalDetails` for error handling instead.

## Security Configuration

### WebSocketSetCertValidation

```vb
Public Sub WebSocketSetCertValidation(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables server certificate chain validation after the TLS handshake. Disabled by default because many WebSocket servers use self-signed or internal-PKI certificates.

When enabled, Wasabi uses `CertGetCertificateChain` and `CertVerifyCertificateChainPolicy` to check that the server's certificate is trusted and its name matches the host.

#### Example

```vb
' Require valid certificate for production
WebSocketSetCertValidation True, h
WebSocketConnect "wss://api.example.com/ws", h
```

### WebSocketSetRevocationCheck

```vb
Public Sub WebSocketSetRevocationCheck(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables CRL/OCSP revocation checking during server certificate validation.

When enabled and `WebSocketSetCertValidation` is also `True`, Wasabi passes `CERT_CHAIN_REVOCATION_CHECK_CHAIN` to `CertGetCertificateChain`. If the CRL or OCSP responder is unreachable, the connection fails with `ERR_CERT_VALIDATE_FAILED`.

Default: disabled (maximizes compatibility with firewalled environments).

#### Example

```vb
' Full certificate validation with revocation check
WebSocketSetCertValidation True, h
WebSocketSetRevocationCheck True, h
WebSocketConnect "wss://api.example.com/ws", h
```

### WebSocketSetClientCert

```vb
Public Sub WebSocketSetClientCert(ByVal thumbprintOrSubject As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Configures a client certificate for TLS mutual authentication (mTLS). The parameter is matched against the `Subject` field of certificates in the Current User Personal store.

### WebSocketSetClientCertPfx

```vb
Public Sub WebSocketSetClientCertPfx(ByVal pfxPath As String, ByVal pfxPassword As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Loads a client certificate from a PFX file and its password. This certificate is used for TLS mutual authentication.

> [!IMPORTANT]
> Client certificate settings must be applied before `WebSocketConnect` and are ignored once the TLS handshake has begun.

### WebSocketSetHttp2

```vb
Public Sub WebSocketSetHttp2(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Requests HTTP/2 during the TLS handshake by advertising the `h2` protocol via ALPN.

> [!NOTE]
> Support for HTTP/2 over WebSocket (RFC 8441) is experimental. Even when
> this flag is enabled, the server must support `h2` ALPN negotiation and
> HTTP/2 WebSocket upgrades.

## MQTT Client

Wasabi includes a minimal MQTT 3.1.1 client that uses the existing WebSocket transport. This allows direct connection to MQTT brokers that support WebSocket listeners (e.g., Mosquitto, HiveMQ, AWS IoT).

All MQTT functions share the same WebSocket connection handle. You must call `WebSocketConnect` with a WebSocket URL before using any MQTT function.

### MqttConnect

```vb
Public Function MqttConnect(ByVal clientId As String, Optional ByVal username As String, Optional ByVal password As String, Optional ByVal keepAlive As Integer = 60, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends an MQTT CONNECT packet over the established WebSocket connection.

#### Example

```vb
WebSocketConnect "wss://broker.emqx.io:8084/mqtt", h
MqttConnect "wasabi-client", h
```

### MqttPublish

```vb
Public Function MqttPublish(ByVal topic As String, ByVal message As String, Optional ByVal qos As Byte = 0, Optional ByVal retained As Boolean = False, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Publishes a text message to the given topic.

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

Polls for incoming MQTT messages. Returns a string in the format `topic|payload` when a PUBLISH packet is received, or an empty string when no message is available.

Internally uses `WebSocketReceiveBinaryCheck` and a state-machine parser to decode MQTT packets.

#### Example

```vb
Dim msg As String
msg = MqttReceive(h)
If msg <> "" Then
    Dim parts() As String
    parts = Split(msg, "|", 2)
    Debug.Print "Topic:", parts(0)
    Debug.Print "Payload:", parts(1)
End If
```

> [!NOTE]
> The MQTT client supports QoS 0 (at most once) for publish. QoS 1 and 2 are not yet implemented.

## Error Reference

### WasabiError Enumeration

The full enumeration now includes codes for certificate loading, certificate validation, fragment overflow, and TLS renegotiation.

```vb
Public Enum WasabiError
    ERR_NONE = 0
    ERR_WSA_STARTUP_FAILED = 1
    ERR_SOCKET_CREATE_FAILED = 2
    ERR_DNS_RESOLVE_FAILED = 3
    ERR_CONNECT_FAILED = 4
    ERR_TLS_ACQUIRE_CREDS_FAILED = 5
    ERR_TLS_HANDSHAKE_FAILED = 6
    ERR_TLS_HANDSHAKE_TIMEOUT = 7
    ERR_WEBSOCKET_HANDSHAKE_FAILED = 8
    ERR_WEBSOCKET_HANDSHAKE_TIMEOUT = 9
    ERR_SEND_FAILED = 10
    ERR_RECV_FAILED = 11
    ERR_NOT_CONNECTED = 12
    ERR_ALREADY_CONNECTED = 13
    ERR_TLS_ENCRYPT_FAILED = 14
    ERR_TLS_DECRYPT_FAILED = 15
    ERR_INVALID_URL = 16
    ERR_HANDSHAKE_REJECTED = 17
    ERR_CONNECTION_LOST = 18
    ERR_INVALID_HANDLE = 19
    ERR_MAX_CONNECTIONS = 20
    ERR_PROXY_CONNECT_FAILED = 21
    ERR_PROXY_AUTH_FAILED = 22
    ERR_PROXY_TUNNEL_FAILED = 23
    ERR_INACTIVITY_TIMEOUT = 24
    ERR_CERT_LOAD_FAILED = 25
    ERR_CERT_VALIDATE_FAILED = 26
    ERR_FRAGMENT_OVERFLOW = 27
    ERR_TLS_RENEGOTIATE = 28
End Enum
```

| Code | Name | Cause |
|:---|:---|:---|
| 0 | `ERR_NONE` | No error |
| 1 | `ERR_WSA_STARTUP_FAILED` | `WSAStartup` returned a non-zero code |
| 2 | `ERR_SOCKET_CREATE_FAILED` | `socket()` returned `INVALID_SOCKET` |
| 3 | `ERR_DNS_RESOLVE_FAILED` | `gethostbyname()` returned null or WSA error 11001–11004 |
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
| 17 | `ERR_HANDSHAKE_REJECTED` | Server returned non‑101 status or `Sec-WebSocket-Accept` was invalid |
| 18 | `ERR_CONNECTION_LOST` | `recv()` returned zero or an oversized frame was received |
| 19 | `ERR_INVALID_HANDLE` | Handle is out of range (reserved for future validation) |
| 20 | `ERR_MAX_CONNECTIONS` | All 64 pool slots are in use |
| 21 | `ERR_PROXY_CONNECT_FAILED` | Could not send CONNECT or proxy did not respond |
| 22 | `ERR_PROXY_AUTH_FAILED` | Proxy returned HTTP 407 (or SOCKS5 auth rejected) |
| 23 | `ERR_PROXY_TUNNEL_FAILED` | Proxy returned a non‑200 status for CONNECT |
| 24 | `ERR_INACTIVITY_TIMEOUT` | No data received within the configured inactivity window |
| 25 | `ERR_CERT_LOAD_FAILED` | Failed to load client certificate from PFX or Windows store |
| 26 | `ERR_CERT_VALIDATE_FAILED` | Server certificate chain validation failed |
| 27 | `ERR_FRAGMENT_OVERFLOW` | Received fragmented message exceeds the fragment buffer size |
| 28 | `ERR_TLS_RENEGOTIATE` | Server requested TLS renegotiation (not supported) |

> [!NOTE]
> `ERR_TLS_RENEGOTIATE` replaces the former reuse of `ERR_TLS_DECRYPT_FAILED` for the `SEC_I_RENEGOTIATE` condition. Connections encountering this error will be closed; auto‑reconnect can recover.

## Practical Patterns

### Single connection workbook

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

### Multi connection workbook

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

    WebSocketSetPingInterval 30000, g_MarketHandle
    WebSocketSetPingInterval 30000, g_EventHandle

    Application.OnTime Now + TimeValue("00:00:01"), "PollConnections"
End Sub

Sub PollConnections()
    Dim market As String
    Dim event As String

    market = WebSocketReceive(g_MarketHandle)
    event = WebSocketReceive(g_EventHandle)

    If market <> "" Then Sheet1.Range("A1").Value = market
    If event <> "" Then Sheet1.Range("B1").Value = event

    If WebSocketIsConnected(g_MarketHandle) Or WebSocketIsConnected(g_EventHandle) Then
        Application.OnTime Now + TimeValue("00:00:01"), "PollConnections"
    End If
End Sub
```

### Error handling pattern

```vb
Sub ConnectWithDiagnostics()
    Dim h As Long

    If WebSocketConnect("wss://echo.websocket.org", h) Then
        Debug.Print "Connected, handle:", h
        Debug.Print "Host:", WebSocketGetHost(h)
        Debug.Print "Port:", WebSocketGetPort(h)
        Debug.Print "Path:", WebSocketGetPath(h)
    Else
        ' Use the new combined description
        Debug.Print WebSocketGetErrorDescription(h)
    End If
End Sub
```

### MQTT IoT dashboard

```vb
Sub StartMqttDashboard()
    Dim h As Long
    If Not WebSocketConnect("wss://test.mosquitto.org:8081/mqtt", h) Then
        Debug.Print "Connection failed"
        Exit Sub
    End If
    MqttConnect "wasabi-dashboard", h
    MqttSubscribe "sensors/temperature", 0, h
    MqttSubscribe "sensors/humidity", 0, h

    Do
        Dim msg As String
        msg = MqttReceive(h)
        If msg <> "" Then
            Dim parts() As String
            parts = Split(msg, "|", 2)
            If parts(0) = "sensors/temperature" Then
                Sheet1.Cells(2, 1).Value = Now()
                Sheet1.Cells(2, 2).Value = parts(1)
            ElseIf parts(0) = "sensors/humidity" Then
                Sheet1.Cells(3, 1).Value = Now()
                Sheet1.Cells(3, 2).Value = parts(1)
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
> Wasabi operates entirely on the VBA main thread. There are no background threads. All socket activity, maintenance, and reconnect logic runs when your code explicitly calls a Wasabi function.

> [!WARNING]
> Automatic heartbeat, ping scheduling, inactivity timeout detection, and auto reconnect triggering are all driven by polling calls. If your code stops calling receive functions, these features stop functioning.

> [!WARNING]
> Queue capacity is fixed at 512 messages per type per connection. Under sustained high message rates, messages will be dropped without error if the queue is not drained fast enough.

> [!CAUTION]
> Custom headers, subprotocol, proxy configuration, buffer sizes, MTU settings, and security options must be set before calling `WebSocketConnect`. Changes after connection have no effect on the active session.

> [!CAUTION]
> `WebSocketDisconnect` always disables auto reconnect. This is intentional and cannot be overridden.

> [!NOTE]
> The pipe-delimited format returned by `WebSocketGetStats`, `WebSocketGetReconnectInfo`, `WebSocketGetProxyInfo`, and `WebSocketGetMTUInfo` is intended for human-readable diagnostics. Do not build parsing logic that depends on field order or format stability across future versions.
