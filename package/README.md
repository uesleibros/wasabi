# package

This folder contains the distributable file for Wasabi.

## What this file is

`Wasabi.bas` is a standard VBA module exported directly from the VBA IDE. It is
a plain text file that the VBA runtime can import, parse and compile natively.

There is no build step, no compilation, no binary, and no packaging involved.
What you see in this file is exactly what runs inside the Office process.

## What this file contains

Internally, `Wasabi.bas` is organized into several distinct layers that work
together to provide a complete WebSocket stack:

**API declarations**
All Windows API functions used by Wasabi are declared at the top of the module
using `Declare PtrSafe` (VBA7/64-bit) or `Declare` (VBA6/32-bit), selected
automatically via `#If VBA7` conditional compilation. This includes functions
from `ws2_32.dll`, `secur32.dll` and `kernel32.dll`.

**Type definitions**
All Win32 structures required by the API calls are defined as VBA `Type` blocks,
including `WSADATA`, `SOCKADDR_IN`, `SecHandle`, `SecBuffer`, `SecBufferDesc`,
`SCHANNEL_CRED`, `SecPkgContext_StreamSizes`, `FD_SET`, `TIMEVAL`, `HOSTENT32`
and `HOSTENT64`. The HOSTENT structure is defined in two versions to handle
pointer size differences between 32-bit and 64-bit environments.

**Connection pool**
Wasabi manages up to 64 simultaneous connections through a statically allocated
array of `WasabiConnection` UDTs. Each entry holds the full state of one
connection: socket handle, TLS context, receive and decrypt buffers, message
queues, fragment buffer, reconnect configuration, proxy settings, statistics and
more. Connections are identified by integer handles returned to the caller.

**TLS stack**
The TLS layer is implemented manually using the Windows SSPI Schannel provider.
This includes the full handshake loop via `InitializeSecurityContext`, stream
encryption via `EncryptMessage`, stream decryption via `DecryptMessage` with
`SECBUFFER_EXTRA` handling for partial records, and `SEC_I_RENEGOTIATE`
detection. The Schannel credential is configured to support TLS 1.2 and TLS 1.3
via `SP_PROT_TLS1_2_CLIENT` and `SP_PROT_TLS1_3_CLIENT`.

**WebSocket framing**
Frame parsing and frame construction are implemented at the bit level according
to RFC 6455. This includes three-tier payload length decoding (7-bit, 16-bit and
64-bit extended), XOR masking with a per-frame random key, opcode routing for
text, binary, continuation, ping, pong and close frames, and reassembly of
fragmented messages via a dedicated fragment buffer per connection.

**SHA-1 implementation**
The WebSocket handshake requires computing `SHA-1(key + GUID)` and encoding the
result in Base64 to produce the `Sec-WebSocket-Accept` header value. Wasabi
implements SHA-1 entirely in VBA without any external dependency, including the
message schedule expansion, the round function selection, the compression
function and the big-endian serialization of the digest. Unsigned 32-bit
arithmetic is emulated using signed `Long` values with bitmask operations to
avoid overflow.

**Base64 encoder**
Used for encoding the `Sec-WebSocket-Key` and the SHA-1 digest. Also used
internally for proxy Basic authentication credentials.

**DNS resolution**
Hostname resolution is handled via `gethostbyname` with manual pointer
arithmetic to extract the first IPv4 address from the returned `HOSTENT`
structure. The pointer size difference between 32-bit and 64-bit is handled by
using two separate `HOSTENT` type definitions and selecting the correct one at
runtime via conditional compilation.

**Non-blocking I/O**
All sockets are switched to non-blocking mode via `ioctlsocket` with `FIONBIO`
immediately after creation. Connection establishment uses `select` with a
10-second timeout on the writable set. Data availability is checked via
`ioctlsocket` with `FIONREAD` before each `recv` call. This ensures that no
call ever blocks the VBA thread.

**Ring buffer queues**
Received text messages and binary messages are stored in separate circular
queues with a fixed capacity of 512 entries each. The queues use head and tail
indices with modular arithmetic for O(1) enqueue and dequeue operations without
any memory allocation or copying beyond the initial buffer setup.

## Portability

This single file is the entire distribution. No other files are required. No
references need to be enabled. No DLLs need to be registered. No installer
needs to be run.
