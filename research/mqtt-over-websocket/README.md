# MQTT over WebSocket Parser

Wasabi includes a built‑in MQTT packet parser that works over
WebSocket connections, enabling MQTT client functionality without
additional libraries.

## Supported features

- CONNECT, CONNACK, PUBLISH, PUBACK, SUBSCRIBE, SUBACK, UNSUBSCRIBE,
  UNSUBACK, PINGREQ, PINGRESP, DISCONNECT.
- QoS 0 and 1 (basic ACK only, no retry).
- Retained flag.
- Username/password authentication.

## Limitations

- No QoS 2.
- No session persistence.
- No will message.
- Packet size limited to 4KB (parser buffer).
