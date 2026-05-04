# MQTT Parser State Machine

Wasabi uses a simple 4‑stage state machine to parse MQTT byte streams.

## Stages

- **Stage 0**: Read the fixed header byte. Extract packet type from
  upper 4 bits (`MqttCurrentPacketType`) and flags from lower 4 bits
  (`MqttCurrentFlags`). Transition to Stage 1.

- **Stage 1**: Read the Remaining Length field (variable‑length encoding).
  Accumulate bytes until the MSB is 0. Compute
  `MqttExpectedRemaining`. Transition to Stage 2.

- **Stage 2**: Read the variable header and payload into `MqttBuffer`
  until `MqttBufLen >= MqttExpectedRemaining`. Transition to Stage 3.

- **Stage 3**: Packet is complete (`MqttHasPacket = True`). The caller
  (e.g., `MqttReceive`) can process it and call `MqttResetParser` to
  return to Stage 0.

## Remaining length encoding

`MqttEncodeRemainingLength` implements the variable‑length encoding
from MQTT spec (each byte uses 7 bits, MSB indicates continuation).

## Integration with WebSocket

In `MqttReceive`, the parser is fed with raw bytes from binary
WebSocket messages. It loops until a complete packet is assembled or
the queue is empty.
