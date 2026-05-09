# Protocol Extensions

Protocol handlers let you intercept WebSocket text and binary messages directly,
bypassing the internal message queue. This is the cleanest way to implement an
application-layer protocol (MQTT 5, AMQP, a custom binary format) without
cluttering your main VBA loop with frame-parsing logic.

## Required interface

Create a VBA class that implements the following public methods:

```vb
' MyProtocol.cls

Public Sub OnConnect(ByVal handle As Long)
    ' Called once, immediately after the WebSocket handshake completes.
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    ' Called when the connection closes, whether by normal close, error, or manual disconnect.
End Sub

Public Sub OnTextMessage(ByVal handle As Long, ByVal message As String)
    ' Delivers every received text frame as a UTF-8 string.
End Sub

Public Sub OnBinaryMessage(ByVal handle As Long, ByRef data() As Byte)
    ' Delivers every received binary frame.
End Sub
```

> [!NOTE]
> When a protocol handler is registered, the default `WebSocketReceive` and
> `WebSocketReceiveBinary` queues are bypassed for that connection. All messages
> must be consumed inside the handler. This gives you zero-copy delivery and full
> control over the data flow.

## Registration

Register per handle with `WasabiUseProtocol`:

```vb
Dim h As Long

If WebSocketConnect("wss://broker.example.com/mqtt", h, , , "mqtt") Then
    Dim proto As New Mqtt5Protocol
    WasabiUseProtocol proto, h
End If
```

You can swap the handler at any time. If the connection is already open, the new
handler receives an `OnConnect` callback immediately.

## Lifecycle example: MQTT 5 client

Below is a simplified sketch of an MQTT 5 protocol handler. A production implementation
would manage session state, packet IDs, and publish/subscribe workflows entirely inside
the class.

```vb
' Mqtt5Protocol.cls

Private m_Handle As Long
Private m_State  As Byte   ' 0 = disconnected, 1 = connecting, 2 = connected

Public Sub OnConnect(ByVal handle As Long)
    m_Handle = handle
    m_State  = 1
    BuildAndSendConnect handle   ' send MQTT CONNECT with MQTT 5 properties
End Sub

Public Sub OnTextMessage(ByVal handle As Long, ByVal message As String)
    ' Text frames are unusual in MQTT; ignore or log.
End Sub

Public Sub OnBinaryMessage(ByVal handle As Long, ByRef data() As Byte)
    ' Parse the MQTT packet type and dispatch:
    '   CONNACK    -> transition to connected, start subscriptions
    '   PUBLISH    -> call user callback, send PUBACK / PUBREC as needed
    '   SUBACK     -> mark subscription confirmed
    '   PUBREC / PUBREL -> update in-flight tracker
    '   DISCONNECT -> handle reason codes
    ParseAndProcess handle, data
End Sub

Public Sub OnDisconnect(ByVal handle As Long)
    m_State = 0
End Sub
```

> [!TIP]
> Your handler runs inside the frame-processing pipeline, so it can call the public
> Wasabi API (for example, `WebSocketSendBinary`) directly from `OnBinaryMessage`
> without re-entering the frame queue.

## Interaction with the engine

The handler is called after all inbound middlewares have run. The `handle` parameter
always identifies the specific connection that triggered the callback. Compression is
also transparent: messages arrive already inflated, regardless of whether
`permessage-deflate` was negotiated.
