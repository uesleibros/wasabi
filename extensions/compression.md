# Compression Extensions

The Wasabi engine decouples compression from the core module. By registering a
compression handler you can replace the built-in `permessage-deflate` (zlib) with
any algorithm (LZ4, Brotli, Zstandard, or a custom enterprise codec) without
modifying `Wasabi.bas`.

## Required interface

Your class must implement the following methods:

```vb
' MyCompressor.cls

Public Sub OnConnect(ByVal handle As Long)
    ' Called when the connection is ready, after the WebSocket handshake.
    ' Use to initialise compressor or decompressor state.
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    ' Called before the connection is terminated.
    ' Release any resources, e.g. streaming context handles.
End Sub

Public Function Deflate(ByRef data() As Byte, _
                        ByVal windowBits As Long, _
                        ByVal contextTakeover As Boolean) As Byte()
    ' Compress the payload.
    ' windowBits:      negotiated window size (negative value per zlib convention)
    ' contextTakeover: True if the previous compression state should be retained
    ' Returns the compressed byte array; may be empty.
End Function

Public Function Inflate(ByRef data() As Byte, _
                        ByVal windowBits As Long, _
                        ByVal contextTakeover As Boolean) As Byte()
    ' Decompress the payload.
    ' Returns the decompressed byte array.
End Function
```

> [!NOTE]
> The engine calls `Deflate` and `Inflate` only when the WebSocket `permessage-deflate`
> extension has been successfully negotiated, or when you manually enable it for a TCP
> connection. If no handler is registered, no compression occurs.

## Registration

```vb
Dim lz4 As New LZ4Compressor
WasabiUseCompression lz4, handle
```

You can register a handler before or after connecting. If the connection is already
open, the handler receives an `OnConnect` callback immediately.

## Integration with `permessage-deflate` negotiation

To use automatic compression negotiation, pass `True` for `DeflateEnabled` in
`WebSocketConnect`. The engine will call your handler instead of the default zlib:

```vb
Dim h   As Long
Dim lz4 As New LZ4Compressor

WebSocketConnect "wss://server/ws", h, True   ' negotiate permessage-deflate
WasabiUseCompression lz4, h
```

The `windowBits` and `contextTakeover` arguments passed to `Deflate` and `Inflate`
reflect the values actually negotiated with the server, so your implementation can
honour them or ignore them as appropriate.

## Example: identity (pass-through) compressor

If you want to disable compression while keeping the negotiation active, for example
during testing, a pass-through handler works:

```vb
' IdentityCompressor.cls

Public Sub OnConnect(ByVal handle As Long):    End Sub
Public Sub OnDisconnect(ByVal handle As Long): End Sub

Public Function Deflate(ByRef data() As Byte, _
                        ByVal windowBits As Long, _
                        ByVal contextTakeover As Boolean) As Byte()
    Deflate = data
End Function

Public Function Inflate(ByRef data() As Byte, _
                        ByVal windowBits As Long, _
                        ByVal contextTakeover As Boolean) As Byte()
    Inflate = data
End Function
```

## Example: LZ4 fast compression

A production LZ4 handler would look like this:

```vb
' LZ4Compressor.cls

Private ctxDeflate As Long
Private ctxInflate As Long

Public Sub OnConnect(ByVal handle As Long)
    ' Allocate LZ4 streaming context if using HC or streaming mode.
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    ' Free streaming context.
End Sub

Public Function Deflate(ByRef data() As Byte, _
                        ByVal windowBits As Long, _
                        ByVal contextTakeover As Boolean) As Byte()
    ' Call LZ4_compress_fast or equivalent and return the compressed frame.
End Function

Public Function Inflate(ByRef data() As Byte, _
                        ByVal windowBits As Long, _
                        ByVal contextTakeover As Boolean) As Byte()
    ' Call LZ4_decompress_safe and return the decompressed frame.
End Function
```

> [!TIP]
> The engine always calls `OnDisconnect` before releasing the connection, giving you
> a guaranteed opportunity to destroy native resources.

## Interaction with middleware and processing order

Middlewares always see uncompressed payloads. The processing order for outbound data is:

1. Middleware `OnBeforeSend`
2. Compression handler `Deflate` (if registered and active)
3. WebSocket framing

For inbound data, the order is reversed:

1. WebSocket deframing
2. Compression handler `Inflate`
3. Middleware `OnAfterReceive`
