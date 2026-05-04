# Buffer Sizing

## Receive buffer (`SO_RCVBUF`)

- **Size:** 256 KB (`BUFFER_SIZE`).
- **Why 256 KB?** VBA may be occupied for several milliseconds processing frames
  or executing `DoEvents`. A large kernel buffer ensures that incoming data is
  not lost while VBA is busy.
- **User override:** `WebSocketSetBufferSizes` allows sizes from 8 KB up to
  16 MB. It cannot be changed while connected.

## Send buffer (`SO_SNDBUF`)

- **Size:** Also 256 KB.
- **Why symmetric?** Simplicity. WebSocket traffic is often bidirectional;
  consuming a bit more memory for the send buffer is harmless.

## Fragmentation buffer

- **Size:** `FRAGMENT_BUFFER_SIZE` (256 KB by default).
- **Purpose:** Holds incomplete WebSocket message fragments until reassembly. If
  a single message exceeds this limit, `ERR_FRAGMENT_OVERFLOW` is raised.

## Application‑layer batching

The batch send functions use a separate 64 KB buffer (`batchBuf`) that is
independent of the socket buffers.
