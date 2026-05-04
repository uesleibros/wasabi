# Compression Flow

## Deflate (outgoing messages)

`DeflatePayload(handle, data, dataLen, outLen)` compresses a payload
before framing.

1. **Window bits**: Uses `DeflateWindowBits` (default `-15`, raw deflate).
   This is set from the negotiated `server_max_window_bits`.

2. **Stream initialization**: If `DeflateContextTakeover` is enabled and
   `DeflateReady` is true, the stream is reused. Otherwise a new stream
   is initialized with `deflateInit2`.

3. **Compression**: Data is compressed with `Z_SYNC_FLUSH` (not
   `Z_FINISH`) so the message boundary is preserved but the stream remains
   usable for the next message.

4. **Cleanup**: If context takeover is off, the stream is ended and
   `DeflateReady = False`. Otherwise the stream state is saved in
   `DeflateStream`.

5. **Output**: The compressed bytes are returned. The caller then builds
   a WebSocket frame with RSV1 set.

## Inflate (incoming messages)

`InflatePayload(handle, data, dataLen, outLen)` decompresses a compressed
payload.

1. **Trailer**: A raw deflate stream must end with the bytes
   `0x00 0x00 0xFF 0xFF`. Wasabi appends this trailer to the incoming
   data before passing to zlib.

2. **Stream handling**: Similar to deflate, respects
   `InflateContextTakeover`.

3. **Error**: If `inflate()` returns anything other than `Z_OK` or
   `Z_STREAM_END`, decompression fails, the stream is destroyed, and
   a close frame with code `1007` is sent.

## Why raw deflate?

Raw deflate (window bits = -15) omits the zlib wrapper (2‑byte header
and 4‑byte Adler‑32 footer). RFC 7692 requires this for
`permessage-deflate`.

## Why Z_SYNC_FLUSH?

`Z_SYNC_FLUSH` inserts a deflate boundary that lets the receiver
decompress the message immediately without waiting for the next deflate
block. This is required because WebSocket messages are independent
application‑level units.
