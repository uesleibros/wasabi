# Fragment Reassembly

## Key state fields

- `FragmentBuffer()` – byte array of size `FRAGMENT_BUFFER_SIZE` (256 KB).
- `FragmentLen` – number of bytes accumulated so far.
- `FragmentOpcode` – opcode of the first frame (`0x1` for TEXT, `0x2` for
  BINARY).
- `Fragmenting` – boolean indicating we are inside a fragmented sequence.
- `FragmentIsCompressed` – set from the RSV1 bit of the first frame.

## Algorithm (inside `ProcessFrames`)

### First frame (TEXT or BINARY, `FIN = 0`)

```
Fragmenting = True
FragmentOpcode = opcode
FragmentIsCompressed = isCompressed
FragmentLen = 0
Copy payload to FragmentBuffer[0..payloadLen-1]
FragmentLen = payloadLen
```


### Continuation frame (CONTINUATION, `FIN = 0`)

- If `FragmentLen + payloadLen > buffer size` → raise `ERR_FRAGMENT_OVERFLOW`.
- Else copy payload to `FragmentBuffer[FragmentLen..]`.
- `FragmentLen += payloadLen`.

### Final frame (CONTINUATION, `FIN = 1`)

- Append payload (same safety check).
- If `FragmentIsCompressed` and `DeflateActive`:
  - Call `InflatePayload` on the complete `FragmentBuffer`.
  - If inflate fails, send a Close frame with code `1007` and disconnect.
  - The decompressed result becomes the final payload.
- Based on `FragmentOpcode`:
  - TEXT → convert UTF‑8 to string and enqueue.
  - BINARY → enqueue as a binary message.
- Reset `Fragmenting = False`, `FragmentLen = 0`.

### Direct final frame (TEXT or BINARY, `FIN = 1`)

- No fragment state involved. If compressed, inflate the single payload
  immediately.

## Why not support interleaved fragments?

RFC 6455 requires frames of a fragmented message to be sent consecutively, with
no other frames of other messages in between. Wasabi assumes that and does not
implement multiplexing of fragments.
