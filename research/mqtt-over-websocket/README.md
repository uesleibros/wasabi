# MQTT over WebSocket Parser

Wasabi includes a built‑in MQTT 3.1.1 packet parser that works over WebSocket connections, enabling native MQTT client functionality in VBA without external dependencies.

## Supported Features

- **Full Control Packet Support**: CONNECT, CONNACK, PUBLISH, PUBACK, PUBREC, PUBREL, PUBCOMP, SUBSCRIBE, SUBACK, UNSUBSCRIBE, UNSUBACK, PINGREQ, PINGRESP, DISCONNECT.
- **QoS 0, 1 and 2**: Fully implemented. QoS 2 (Exactly Once) ensures that messages are delivered exactly one time through a four-step handshake, preventing duplicates even in unstable network conditions.
- **Full In-Flight Management**: Both QoS 1 and QoS 2 utilize an internal `MqttInFlight` queue that tracks Packet IDs and clears messages only upon the successful completion of the respective acknowledgment protocol (`PUBACK` for QoS 1; `PUBCOMP` for QoS 2).
- **Dynamic Payload Handling**: The parser buffer uses an elastic allocation strategy (Chunk Allocation), allowing it to handle packets of any size (e.g., large JSON payloads), limited only by system memory.
- **UTF-8 Integrity**: Topic and Message length calculations are based on byte-count rather than character-count, ensuring full compatibility with accented characters, multi-byte strings, and emojis.
- **Retained Messages**: Full support for the retained flag on outgoing and incoming publications.
- **Authentication**: Supports standard MQTT username and password fields.

## Implementation Details

### QoS 2 "Exactly Once" Handshake
When a message is published or received with QoS 2, Wasabi manages the complete state machine:
1. **PUBLISH**: Message is sent and stored in the in-flight queue.
2. **PUBREC**: Broker confirms receipt; Wasabi automatically responds with `PUBREL`.
3. **PUBREL**: Confirmation that the packet ID is released.
4. **PUBCOMP**: Final step; Wasabi automatically responds to incoming `PUBREL` packets and clears its own in-flight messages only when this final ACK arrives.

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

- **Manual Retry**: While messages are tracked In-Flight, the client does not currently perform automatic re-transmission if an acknowledgment is not received within a specific timeout (must be handled by application logic if required).
- **No Last Will and Testament (LWT)**: Will message parameters are not currently exposed in the connection interface.
- **Volatile Session**: Session state is stored in RAM; In-Flight queues are cleared if the VBA project is reset or the workbook is closed.
