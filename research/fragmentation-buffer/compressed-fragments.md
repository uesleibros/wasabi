# Compressed Fragments

When `permessage-deflate` is active, a fragmented message may have the RSV1 bit
set on the **first** frame. All continuation frames must have RSV1 = 0.

## Wasabi’s implementation

- `FragmentIsCompressed` is set from the RSV1 of the first frame.
- On the final frame, if `FragmentIsCompressed = True`, the entire reassembled
  buffer is passed through `InflatePayload`.
- The inflate logic appends the mandatory raw deflate trailer (`0x00 0x00 0xFF
  0xFF`) before decompression.
- If decompression fails, the connection is closed with `1007` (Invalid frame
  payload data).

## Why not decompress each fragment individually?

RFC 7692 requires that all fragments of a message are compressed together as a
single deflate block. Decompressing individually would produce garbage. Wasabi
correctly accumulates the fragments and decompresses once.
