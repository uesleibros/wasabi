# API Reference

This document describes the public API exposed by `Wasabi.bas`.

Wasabi is a native WebSocket and WSS client for VBA built directly on Winsock and Schannel. The public API is intentionally simple, but several runtime behaviors are important to understand before using it in production.

## Core Concepts

### Handles

Every successful connection is represented by an integer handle.

A handle identifies one slot in Wasabi's internal connection pool. Most public functions accept an optional `handle` parameter. If omitted, Wasabi uses the current default handle.

By design, `WebSocketConnect` automatically sets the newly created connection as the default handle.

> [!TIP]
> If your workbook uses only one connection, you can omit the handle parameter in most calls.

### Polling model

Wasabi runs inside VBA, which is single-threaded. Incoming messages are not delivered through background events. Instead, they are buffered internally and exposed through polling functions such as `WebSocketReceive` and `WebSocketReceiveBinary`.

> [!IMPORTANT]
> Your code must call the receive functions regularly if low latency matters.

### Queue model

Wasabi stores received messages in internal queues.

* Text queue capacity: 512 messages
* Binary queue capacity: 512 messages

> [!WARNING]
> If your code does not drain the queues quickly enough, new messages may be dropped when capacity is reached.

## Typical Usage Pattern

```vb
Sub Example()
    Dim h As Long
    Dim ok As Boolean
    Dim msg As String

    ok = WebSocketConnect("wss://echo.websocket.org", h)
    If Not ok Then
        Debug.Print "Connect failed"
        Debug.Print WebSocketGetLastError()
        Debug.Print WebSocketGetTechnicalDetails()
        Exit Sub
    End If

    WebSocketSend "Hello from Wasabi", h

    Do
        msg = WebSocketReceive(h)
        DoEvents
    Loop Until msg <> ""

    Debug.Print "Received:", msg

    WebSocketDisconnect h
End Sub
```

## Connection Management

### WebSocketConnect

```vb
Public Function WebSocketConnect(ByVal url As String, Optional ByRef outHandle As Long = -1) As Boolean
```

Opens a new WebSocket connection.

This function performs the full connection sequence:

* Initializes Winsock if needed
* Allocates a slot in the internal pool
* Parses the URL
* Resolves DNS
* Opens a TCP socket
* Applies socket options
* Establishes a proxy tunnel if configured
* Performs the TLS handshake for `wss://`
* Performs the WebSocket upgrade handshake
* Assigns the connection as the default handle

#### Parameters

* `url`
  * A `ws://` or `wss://` URL
* `outHandle`
  * Receives the allocated connection handle on success

#### Returns

* `True` on success
* `False` on failure

#### Example

```vb
Dim h As Long

If WebSocketConnect("wss://echo.websocket.org", h) Then
    Debug.Print "Connected:", h
Else
    Debug.Print "Connection failed"
    Debug.Print WebSocketGetLastError()
    Debug.Print WebSocketGetTechnicalDetails()
End If
```

> [!WARNING]
> Do not assume that handle `0` is invalid. `0` can be a valid connection handle.

> [!IMPORTANT]
> If you need custom headers, proxy settings, subprotocol, or custom buffer sizes, configure them before calling `WebSocketConnect`.

### WebSocketDisconnect

```vb
Public Sub WebSocketDisconnect(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Closes a connection, sends a normal Close frame if connected, disables auto reconnect, releases native resources, and clears internal state.

#### Example

```vb
WebSocketDisconnect h
```

> [!NOTE]
> If this was the last active connection, Wasabi also releases Winsock resources.

### WebSocketDisconnectAll

```vb
Public Sub WebSocketDisconnectAll()
```

Closes every active connection in the pool.

#### Example

```vb
WebSocketDisconnectAll
```

### WebSocketIsConnected

```vb
Public Function WebSocketIsConnected(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Returns `True` if the specified connection is active.

#### Example

```vb
If WebSocketIsConnected(h) Then
    Debug.Print "Still connected"
End If
```

### WebSocketSendClose

```vb
Public Function WebSocketSendClose(Optional ByVal code As Integer = 1000, Optional ByVal reason As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a WebSocket Close frame with an optional status code and reason string.

#### Example

```vb
WebSocketSendClose 1000, "Normal shutdown", h
```

## Sending Data

### WebSocketSend

```vb
Public Function WebSocketSend(ByVal message As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a UTF-8 text message as a WebSocket text frame.

#### Example

```vb
If Not WebSocketSend("Hello world", h) Then
    Debug.Print "Send failed"
End If
```

> [!NOTE]
> Client frames are masked automatically, as required by RFC 6455.

> [!CAUTION]
> Sending on a disconnected handle returns `False`.

### WebSocketSendBinary

```vb
Public Function WebSocketSendBinary(ByRef data() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends raw binary data as a WebSocket binary frame.

#### Example

```vb
Dim payload(0 To 3) As Byte

payload(0) = &HDE
payload(1) = &HAD
payload(2) = &HBE
payload(3) = &HEF

If WebSocketSendBinary(payload, h) Then
    Debug.Print "Binary sent"
End If
```

### WebSocketBroadcast

```vb
Public Function WebSocketBroadcast(ByVal message As String) As Long
```

Sends the same text message to all connected handles.

#### Returns

The number of successful sends.

#### Example

```vb
Dim delivered As Long
delivered = WebSocketBroadcast("System notification")
Debug.Print "Delivered to", delivered, "connections"
```

### WebSocketBroadcastBinary

```vb
Public Function WebSocketBroadcastBinary(ByRef data() As Byte) As Long
```

Sends the same binary payload to all connected handles.

## Receiving Data

### WebSocketReceive

```vb
Public Function WebSocketReceive(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the next available text message from the queue. If no message is available, returns an empty string.

This function also drives internal maintenance, including:

* socket polling
* TLS decryption
* frame processing
* auto reconnect checks
* inactivity timeout checks
* automatic ping scheduling

#### Example

```vb
Dim msg As String

msg = WebSocketReceive(h)
If msg <> "" Then
    Debug.Print "Received:", msg
End If
```

> [!TIP]
> Call `WebSocketReceive` regularly in loops that use `DoEvents` if you want low latency without freezing the Office UI.

> [!WARNING]
> An empty string can mean either no message is available or the server actually sent an empty payload. If empty payloads matter in your protocol, design accordingly.

### WebSocketReceiveAll

```vb
Public Function WebSocketReceiveAll(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String()
```

Drains the entire text queue and returns all pending messages as a string array.

#### Example

```vb
Dim messages() As String
Dim i As Long

messages = WebSocketReceiveAll(h)

For i = LBound(messages) To UBound(messages)
    Debug.Print messages(i)
Next i
```

> [!CAUTION]
> If the queue is empty, the returned array may not be safe for naive iteration. Defensive bounds handling is recommended.

### WebSocketReceiveBinary

```vb
Public Function WebSocketReceiveBinary(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Byte()
```

Returns the next pending binary message.

#### Example

```vb
Dim data() As Byte

data = WebSocketReceiveBinary(h)

If Not Not data Then
    Debug.Print "Binary received"
End If
```

### WebSocketReceiveBinaryCheck

```vb
Public Function WebSocketReceiveBinaryCheck(ByRef outData() As Byte, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Writes the next binary message into `outData` and returns `True` if data was available.

#### Example

```vb
Dim data() As Byte

If WebSocketReceiveBinaryCheck(data, h) Then
    Debug.Print "Received", UBound(data) - LBound(data) + 1, "bytes"
End If
```

> [!TIP]
> Prefer `WebSocketReceiveBinaryCheck` when you want explicit success and failure semantics.

## Queue Inspection and Control

### WebSocketGetPendingCount

```vb
Public Function WebSocketGetPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the number of queued text messages.

### WebSocketGetBinaryPendingCount

```vb
Public Function WebSocketGetBinaryPendingCount(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the number of queued binary messages.

### WebSocketGetQueueCapacity

```vb
Public Function WebSocketGetQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the remaining capacity of the text queue.

### WebSocketGetBinaryQueueCapacity

```vb
Public Function WebSocketGetBinaryQueueCapacity(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the remaining capacity of the binary queue.

### WebSocketPeek

```vb
Public Function WebSocketPeek(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the next queued text message without removing it.

#### Example

```vb
Debug.Print WebSocketPeek(h)
```

### WebSocketFlushQueue

```vb
Public Sub WebSocketFlushQueue(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Clears both text and binary queues for the specified connection.

#### Example

```vb
WebSocketFlushQueue h
```

> [!WARNING]
> This discards queued messages immediately.

## Control Frames and Heartbeat

### WebSocketSendPing

```vb
Public Function WebSocketSendPing(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a Ping control frame with an optional payload.

#### Example

```vb
WebSocketSendPing "heartbeat", h
```

### WebSocketSendPong

```vb
Public Function WebSocketSendPong(Optional ByVal payload As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Sends a Pong control frame with an optional payload.

#### Example

```vb
WebSocketSendPong "manual pong", h
```

> [!NOTE]
> Server Ping frames are answered automatically by Wasabi. Manual Pong use is generally unnecessary unless your application protocol requires it.

### WebSocketSetPingInterval

```vb
Public Sub WebSocketSetPingInterval(ByVal intervalMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the automatic Ping interval in milliseconds. Use `0` to disable.

#### Example

```vb
WebSocketSetPingInterval 30000, h
```

> [!IMPORTANT]
> Automatic pings are processed during receive and maintenance paths. If your code stops polling entirely, scheduled pings also stop.

## Reconnect and Reliability

### WebSocketSetAutoReconnect

```vb
Public Sub WebSocketSetAutoReconnect(ByVal enabled As Boolean, Optional ByVal maxAttempts As Long = DEFAULT_RECONNECT_MAX_ATTEMPTS, Optional ByVal baseDelayMs As Long = DEFAULT_RECONNECT_BASE_DELAY_MS, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables automatic reconnection with exponential backoff.

#### Example

```vb
WebSocketSetAutoReconnect True, 10, 1000, h
```

#### Behavior

If a session is lost during a receive or maintenance path, Wasabi attempts to reconnect using the original URL and previously configured connection settings.

The reconnect delay grows exponentially until it reaches the configured cap.

#### Preserved settings

* original URL
* proxy settings
* custom headers
* subprotocol
* NoDelay flag
* log callback
* error dialog setting
* ping interval
* receive timeout
* inactivity timeout

> [!WARNING]
> Reconnect attempts run on the same VBA thread and use `DoEvents` while waiting.

> [!CAUTION]
> Calling `WebSocketDisconnect` disables auto reconnect for that handle by design.

### WebSocketGetReconnectInfo

```vb
Public Function WebSocketGetReconnectInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited summary of reconnect state.

#### Example result

```text
AutoReconnect=1|Attempts=2|MaxAttempts=5|BaseDelayMs=1000
```

## Buffer and Performance Configuration

### WebSocketSetBufferSizes

```vb
Public Sub WebSocketSetBufferSizes(ByVal bufferSize As Long, ByVal fragmentSize As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the internal receive buffer size and fragment buffer size for the connection.

#### Example

```vb
WebSocketSetBufferSizes 524288, 524288, h
```

#### Limits

* Receive buffer: 8192 to 16777216 bytes
* Fragment buffer: 4096 to 16777216 bytes

> [!IMPORTANT]
> This function must be called before `WebSocketConnect`.

> [!WARNING]
> Calling this function after the connection is established does not reallocate active buffers.

### WebSocketSetNoDelay

```vb
Public Function WebSocketSetNoDelay(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Boolean
```

Enables or disables `TCP_NODELAY`.

#### Example

```vb
Call WebSocketSetNoDelay(True, h)
```

#### Behavior

If called before connect, the preference is stored and applied when the socket is created. If called after connect, it is applied immediately.

## Proxy Configuration

### WebSocketSetProxy

```vb
Public Sub WebSocketSetProxy(ByVal proxyHost As String, ByVal proxyPort As Long, Optional ByVal proxyUser As String = "", Optional ByVal proxyPass As String = "", Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Configures an HTTP CONNECT proxy for the specified connection.

#### Example without authentication

```vb
WebSocketSetProxy "proxy.company.local", 8080, , , h
```

#### Example with authentication

```vb
WebSocketSetProxy "proxy.company.local", 8080, "myuser", "mypassword", h
```

> [!IMPORTANT]
> Proxy settings must be defined before `WebSocketConnect`.

> [!WARNING]
> If the proxy returns HTTP 407, Wasabi reports `ERR_PROXY_AUTH_FAILED`.

### WebSocketClearProxy

```vb
Public Sub WebSocketClearProxy(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes proxy settings from the specified connection.

### WebSocketGetProxyInfo

```vb
Public Function WebSocketGetProxyInfo(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a summary string describing proxy configuration.

#### Example result

```text
Host=proxy.company.local|Port=8080|Auth=Yes
```

## Handshake Customization

### WebSocketAddHeader

```vb
Public Sub WebSocketAddHeader(ByVal headerName As String, ByVal headerValue As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Adds a custom HTTP header to the upgrade request.

#### Example

```vb
WebSocketAddHeader "Authorization", "Bearer my-token", h
WebSocketAddHeader "X-Client", "ExcelBot", h
```

> [!IMPORTANT]
> Custom headers must be configured before `WebSocketConnect`.

### WebSocketClearHeaders

```vb
Public Sub WebSocketClearHeaders(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Removes all custom handshake headers for the specified connection.

### WebSocketSetSubProtocol

```vb
Public Sub WebSocketSetSubProtocol(ByVal protocol As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the value sent in the `Sec-WebSocket-Protocol` header.

#### Example

```vb
WebSocketSetSubProtocol "graphql-transport-ws", h
```

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

Sets the timeout used by internal `select()` operations.

#### Example

```vb
WebSocketSetReceiveTimeout 10000, h
```

### WebSocketSetInactivityTimeout

```vb
Public Sub WebSocketSetInactivityTimeout(ByVal timeoutMs As Long, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Sets the maximum allowed period without receiving data before the connection is treated as stale.

#### Example

```vb
WebSocketSetInactivityTimeout 60000, h
```

## Diagnostics and State Inspection

### WebSocketGetLastError

```vb
Public Function WebSocketGetLastError(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As WasabiError
```

Returns the most recent Wasabi error code.

### WebSocketGetLastErrorCode

```vb
Public Function WebSocketGetLastErrorCode(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the most recent native error code such as a WSA or SSPI value.

### WebSocketGetTechnicalDetails

```vb
Public Function WebSocketGetTechnicalDetails(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a human-readable technical explanation of the latest error.

#### Example

```vb
If Not WebSocketConnect("wss://bad.host.example", h) Then
    Debug.Print WebSocketGetLastError(h)
    Debug.Print WebSocketGetLastErrorCode(h)
    Debug.Print WebSocketGetTechnicalDetails(h)
End If
```

### WebSocketGetStats

```vb
Public Function WebSocketGetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns a pipe-delimited metrics snapshot.

#### Example result

```text
BytesSent=1520|BytesReceived=2048|MessagesSent=3|MessagesReceived=4|UptimeSeconds=12|Queued=0|BinaryQueued=0|NoDelay=1|Proxy=none
```

#### Example

```vb
Debug.Print WebSocketGetStats(h)
```

### WebSocketResetStats

```vb
Public Sub WebSocketResetStats(Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Resets byte counters, message counters, and connection start time.

### WebSocketGetUptime

```vb
Public Function WebSocketGetUptime(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the uptime of the connection in seconds.

### WebSocketGetHost

```vb
Public Function WebSocketGetHost(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the host component of the connected URL.

### WebSocketGetPort

```vb
Public Function WebSocketGetPort(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As Long
```

Returns the port component of the connected URL.

### WebSocketGetPath

```vb
Public Function WebSocketGetPath(Optional ByVal handle As Long = INVALID_CONN_HANDLE) As String
```

Returns the path component of the connected URL.

## Logging and User Feedback

### WebSocketSetLogCallback

```vb
Public Sub WebSocketSetLogCallback(ByVal callbackName As String, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Registers a VBA macro that receives internal log messages.

#### Example

```vb
Sub WasabiLogger(ByVal msg As String)
    Debug.Print "[WASABI]", msg
End Sub

Sub Setup()
    Dim h As Long
    WebSocketConnect "wss://echo.websocket.org", h
    WebSocketSetLogCallback "WasabiLogger", h
End Sub
```

> [!WARNING]
> The callback must be callable through `Application.Run` and must accept exactly one string parameter.

### WebSocketSetErrorDialog

```vb
Public Sub WebSocketSetErrorDialog(ByVal enabled As Boolean, Optional ByVal handle As Long = INVALID_CONN_HANDLE)
```

Enables or disables interactive error dialogs for the specified connection.

#### Example

```vb
WebSocketSetErrorDialog True, h
```

> [!TIP]
> In unattended automation, leave dialogs disabled and rely on diagnostics functions instead.

## Error Enumeration

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
End Enum
```

## Practical Recommendations

### Single connection workbooks

If your workbook uses only one connection, relying on the default handle is usually acceptable.

```vb
If WebSocketConnect("wss://echo.websocket.org") Then
    WebSocketSend "Hello"
    Debug.Print WebSocketReceive()
    WebSocketDisconnect
End If
```

### Multi connection workbooks

In multi connection scenarios, always preserve and pass handles explicitly.

```vb
Dim marketHandle As Long
Dim chatHandle As Long

WebSocketConnect "wss://market.example/ws", marketHandle
WebSocketConnect "wss://chat.example/ws", chatHandle

WebSocketSend "subscribe", marketHandle
WebSocketSend "join", chatHandle
```

### Polling loops

Keep the UI responsive in continuous loops.

```vb
Do While WebSocketIsConnected(h)
    Dim msg As String
    msg = WebSocketReceive(h)

    If msg <> "" Then
        Debug.Print msg
    End If

    DoEvents
Loop
```

> [!IMPORTANT]
> In Wasabi, polling is not an implementation detail you can ignore. It is part of the execution model.

## Operational Caveats

> [!WARNING]
> Wasabi is single-threaded because VBA is single-threaded.

> [!WARNING]
> Automatic heartbeat and several maintenance operations are driven by polling paths, not by background timers.

> [!WARNING]
> Queue capacity is finite and can be exhausted under sustained high message volume.

> [!CAUTION]
> Configure proxy settings, headers, subprotocol, and custom buffers before calling `WebSocketConnect`.

> [!CAUTION]
> `WebSocketDisconnect` intentionally disables auto reconnect for the target handle.

> [!NOTE]
> `WebSocketGetStats` is intended for diagnostics and observability. It should not be treated as a strict machine-contract format across all future versions.
