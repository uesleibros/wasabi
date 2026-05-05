# MQTT over WebSocket Parser

Wasabi includes a built‑in MQTT 3.1.1 packet parser that works over WebSocket connections, enabling native MQTT client functionality in VBA without external dependencies.

## Supported Features

- **Full Control Packet Support**: CONNECT, CONNACK, PUBLISH, PUBACK, SUBSCRIBE, SUBACK, UNSUBSCRIBE, UNSUBACK, PINGREQ, PINGRESP, DISCONNECT.
- **QoS 0 and 1**: Fully implemented. QoS 1 (At Least Once) includes an internal In-Flight queue that tracks Packet IDs and clears messages only upon receiving `PUBACK` from the broker.
- **Dynamic Payload Handling**: The parser buffer uses an elastic allocation strategy (Chunk Allocation), allowing it to handle packets of any size (e.g., large JSON payloads), limited only by system memory.
- **UTF-8 Integrity**: Tópico and Message length calculations are based on byte-count rather than character-count, ensuring full compatibility with accented characters, multi-byte strings, and emojis.
- **Retained Messages**: Full support for the retained flag on outgoing and incoming publications.
- **Authentication**: Supports standard MQTT username and password fields.

## Implementation Details

### QoS 1 In-Flight Management
When a message is published with QoS 1, Wasabi:
1. Generates a unique 16-bit Packet Identifier.
2. Persists the message in the `MqttInFlight` queue within the connection pool.
3. Automatically clears the message when the corresponding `PUBACK` is parsed during `MqttReceive`.

### Subprotocol Negotiation
The parser requires the WebSocket handshake to explicitly request the `"mqtt"` subprotocol. This is handled by passing the protocol string to `WebSocketConnect`.

### System Acknowledgments
The `MqttReceive` function returns control tags (`[CONNACK]`, `[SUBACK]`, `[UNSUBACK]`) to allow the calling code to synchronize its state with the broker's responses.

## Current Limitations

- **No QoS 2**: Packets requiring the four-way handshake (`PUBREC`, `PUBREL`, `PUBCOMP`) are ignored.
- **Manual Retry**: While messages are tracked In-Flight, the client does not currently perform automatic re-transmission if a `PUBACK` is not received (must be handled by application logic if required).
- **No Last Will and Testament (LWT)**: Will message parameters are not currently exposed in the connection interface.
- **Volatile Session**: Session state is stored in RAM; In-Flight queues are cleared if the VBA project is reset or the workbook is closed.
