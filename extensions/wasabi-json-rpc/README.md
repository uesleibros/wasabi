# WasabiJsonRpc

A JSON-RPC 2.0 protocol handler for Wasabi. Import `WasabiJsonRpc.cls` alongside
`Wasabi.bas` and register it on any WebSocket connection. No external libraries or
COM references are required.

## Import

1. In the VBA editor, click **File -> Import File...**
2. Import `WasabiJsonRpc.cls`.

The class will appear in your project tree alongside your other modules.

## Quick start

```vb
Dim rpc As New WasabiJsonRpc
Dim h   As Long

' Connect and register the handler.
WebSocketConnect "wss://api.example.com/rpc", h
WasabiUseProtocol rpc, h

' Send a request. The returned id identifies the response.
Dim id As Long
id = rpc.Request("subtract", "{""a"":10,""b"":3}", h)

' Block until the response arrives (up to 3 seconds).
Dim result As String
result = rpc.Receive(id, timeoutMs:=3000, handle:=h)

Debug.Print result   ' e.g. 7

WebSocketDisconnect h
```

## Sending

**`Request(method, params, handle) As Long`**
Sends a JSON-RPC request and returns the auto-assigned integer id. `params` must be
a valid JSON value: an object (`{...}`), an array (`[...]`), or `null`.

**`Notify(method, params, handle)`**
Sends a notification. No response is expected and no id is assigned.

**`Batch(methods(), params(), handle) As Long()`**
Sends multiple requests in a single frame. `methods` and `params` are parallel
string arrays. Returns an array of ids in the same order.

```vb
Dim methods(1) As String
Dim params(1)  As String
Dim ids()      As Long

methods(0) = "add"      : params(0) = "{""a"":1,""b"":2}"
methods(1) = "multiply" : params(1) = "{""a"":3,""b"":4}"

ids = rpc.Batch(methods, params, h)
```

## Receiving (polling)

**`Receive(id, timeoutMs, handle) As String`**
Blocks until the response for `id` arrives or the timeout expires. Returns the raw
JSON result value on success, or an empty string on timeout or RPC error.

**`ReceiveError(id, outCode, outMessage) As Boolean`**
After `Receive` returns empty, call this to distinguish a timeout from an RPC error.
Returns `True` and populates `outCode` and `outMessage` if the server returned a
JSON-RPC error object.

```vb
Dim result As String
result = rpc.Receive(id, 5000, h)

If result = "" Then
    Dim code As Long, msg As String
    If rpc.ReceiveError(id, code, msg) Then
        Debug.Print "RPC error " & code & ": " & msg
    Else
        Debug.Print "Timeout"
    End If
End If
```

**`ReceivePending(outId, outResult, outErrCode, outErrMsg) As Boolean`**
Non-blocking. Returns `True` and populates the out parameters if any unread
response is waiting in the queue. Useful inside a `DoEvents` loop.

**`PendingCount() As Long`**
Returns the number of unread responses currently in the internal queue.

## Receiving (callbacks)

Call `SetCallbackModule` with the name of any standard VBA module. Whenever a
response or notification arrives, the handler will call the corresponding
procedure on that module via `Application.Run`.

```vb
rpc.SetCallbackModule "RpcCallbacks"
```

Your callback module must expose any of the following procedures it wants to handle:

```vb
' RpcCallbacks.bas

Public Sub RpcOnResult(ByVal id As Long, ByVal result As String, ByVal handle As Long)
    Debug.Print "Result for " & id & ": " & result
End Sub

Public Sub RpcOnError(ByVal id As Long, ByVal code As Long, ByVal message As String, ByVal handle As Long)
    Debug.Print "Error " & code & " for " & id & ": " & message
End Sub

Public Sub RpcOnNotify(ByVal method As String, ByVal params As String, ByVal handle As Long)
    Debug.Print "Notification [" & method & "]: " & params
End Sub
```

You do not need to implement all three. Unimplemented procedures are silently skipped.
Callbacks and polling can be used together on the same connection.

## Notifications from the server

Server-sent notifications (messages without an `id` field) are delivered exclusively
through `RpcOnNotify`. They are not placed in the polling queue.

## JSON params format

Params must be passed as a raw JSON string. Wasabi does not include a JSON builder,
so construct the string yourself or use any JSON module available in your project.

```vb
' Object params
rpc.Request "greet", "{""name"":""Alice""}", h

' Array params
rpc.Request "sum", "[1, 2, 3]", h

' No params
rpc.Request "ping", "null", h
```

## Notes

The internal response queue holds up to 256 unread entries. When the queue is full,
the oldest entry is dropped to make room for the incoming one. For high-throughput
workloads, consume responses promptly or increase `MAX_QUEUE` in the class source.

The JSON parser inside `WasabiJsonRpc` is intentionally minimal: it handles the
top-level structure of JSON-RPC 2.0 envelopes correctly but does not attempt to
parse nested application payloads. The `result` and `params` fields are returned
as raw JSON strings for your own code to interpret.
