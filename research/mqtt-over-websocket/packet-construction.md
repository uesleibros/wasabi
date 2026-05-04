# Building MQTT Packets

## CONNECT Packet

`BuildMqttConnectPacket(clientId, username, password, keepAlive)`
constructs a CONNECT packet following MQTT 3.1.1:

- Variable header: Protocol Name ("MQTT"), Protocol Level (4),
  Connect Flags, Keep Alive.
- Payload: Client ID, optional username, optional password.

## Other Packets

`MqttBuildPacket(ptype, flags, payload, payloadLen)` builds a generic
MQTT packet:

- Fixed header: `ptype * 16 | flags` (first byte), followed by
  encoded remaining length.
- Payload is copied as‑is.

## UTF‑8 Strings

MQTT requires UTF‑8 encoding for string fields. Wasabi uses the same
`StringToUtf8` function used for WebSocket text frames.
