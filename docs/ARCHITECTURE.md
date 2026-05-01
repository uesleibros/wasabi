# Architecture

This document describes the internal architecture of Wasabi. It is intended
for contributors, advanced users, and anyone who wants to understand how the
module works under the hood.

## High-Level Overview

Wasabi is a single-file VBA module that implements a complete WebSocket client
stack using only native Windows APIs. It does not depend on any external library,
COM component, or registered DLL.

The module is organized into five distinct layers, each responsible for one
aspect of the communication pipeline.

```mermaid
graph TD
    A["Your VBA Code<br/>WebSocketConnect / Send / Receive"]
    B["Public API Layer<br/>Handle resolution, validation, routing"]
    C["WebSocket Protocol Layer<br/>Frame construction, parsing, masking,<br/>fragmentation, control frame handling"]
    D["TLS Layer (Schannel)<br/>Handshake, EncryptMessage, DecryptMessage"]
    E["Transport Layer (Winsock)<br/>socket, connect, send, recv, select"]
    F["Windows Kernel<br/>TCP/IP stack, network driver"]

    A --> B --> C --> D --> E --> F

    style A fill:#f9f,stroke:#333,stroke-width:2px
    style B fill:#ccf,stroke:#333,stroke-width:2px
    style C fill:#cfc,stroke:#333,stroke-width:2px
    style D fill:#ffc,stroke:#333,stroke-width:2px
    style E fill:#fca,stroke:#333,stroke-width:2px
    style F fill:#ddd,stroke:#333,stroke-width:2px
```

## Connection Pool

Wasabi manages all connections through a statically allocated pool of 64
`WasabiConnection` entries. Each entry holds the complete state of one
WebSocket session.

### Structure

The pool is an array of `WasabiConnection` user-defined types, initialized
on the first call to any Wasabi function via `InitConnectionPool`.

```mermaid
graph LR
    subgraph Pool["Connection Pool (64 slots)"]
        direction LR
        A0[0]:::active --- A1[1]:::inactive --- A2[2]:::active --- A3[3]:::inactive --- DOTS[...] --- A62[62]:::inactive --- A63[63]:::active
    end
    subgraph Legend
        L1[● Active]:::active --- L2[○ Available]:::inactive
    end
    classDef active fill:#2ecc71,stroke:#27ae60,color:#fff
    classDef inactive fill:#bdc3c7,stroke:#95a5a6,color:#333
    classDef dots fill:none,stroke:none,color:#333
    class DOTS dots
    style Pool fill:none,stroke:#3498db,stroke-width:2px
    style Legend fill:none,stroke:none
```

### Allocation

When `WebSocketConnect` is called, `AllocConnection` scans the pool for the
first slot where `Connected = False` and `Socket = INVALID_SOCKET`. It
initializes all fields to their defaults and returns the index as the
connection handle.

### Deallocation

When `WebSocketDisconnect` is called, `CleanupHandle` closes the socket,
releases TLS resources, and resets all fields in the slot. The slot becomes
available for reuse.

### Handle Resolution

Most public functions accept an optional handle parameter. The internal
function `ResolveHandle` translates this.

```
If handle = INVALID_CONN_HANDLE (-1)
    → use m_DefaultHandle
Else
    → use the provided handle directly
```

> [!NOTE]
> The default handle is updated automatically by `WebSocketConnect` to point
> to the most recently opened connection.

### Per-Connection State

Each `WasabiConnection` entry contains:

| Category | Fields |
|:---|:---|
| Socket | `Connected`, `TLS`, `Host`, `Port`, `Path` |
| TLS | `hCred`, `hContext`, `Sizes` |
| Receive | `RecvBuffer()`, `RecvLen`, `DecryptBuffer()`, `DecryptLen` |
| Text Queue | `MsgQueue()`, `MsgHead`, `MsgTail`, `MsgCount` |
| Binary Queue | `BinaryQueue()`, `BinaryHead`, `BinaryTail`, `BinaryCount` |
| Fragmentation | `FragmentBuffer()`, `FragmentLen`, `FragmentOpcode`, `Fragmenting` |
| Reconnect | `AutoReconnect`, `ReconnectMaxAttempts`, `ReconnectAttempts`, `ReconnectBaseDelayMs` |
| Proxy | `ProxyHost`, `ProxyPort`, `ProxyUser`, `ProxyPass`, `ProxyEnabled`, `ProxyType` |
| Heartbeat | `PingIntervalMs`, `LastPingSentAt` |
| Timeouts | `ReceiveTimeoutMs`, `InactivityTimeoutMs`, `LastActivityAt` |
| Headers | `CustomHeaders()`, `CustomHeaderCount`, `SubProtocol` |
| Statistics | `Stats` (BytesSent, BytesReceived, MessagesSent, MessagesReceived, ConnectedAt) |
| Diagnostics | `LastError`, `LastErrorCode`, `TechnicalDetails` |
| Logging | `LogCallback`, `EnableErrorDialog` |
| Configuration | `NoDelay`, `CustomBufferSize`, `CustomFragmentSize`, `OriginalUrl` |

## Connection Sequence

The full connection sequence is handled by the internal `ConnectHandle`
function. Every connection, including reconnections, passes through this
same path.

```mermaid
graph TD
    A[WebSocketConnect]
    B["ParseURL<br/>extract host, port, path, scheme"]
    C["ResolveAndConnect<br/>getaddrinfo → Happy Eyeballs → TCP connect"]
    D["ApplySocketOptions<br/>TCP_NODELAY, SO_KEEPALIVE, buffer sizes"]
    E{ProxyEnabled?}
    F["DoProxyHTTP or DoProxySOCKS5"]
    G{TLS wss://?}
    H["AcquireCredentialsHandle<br/>DoTLSHandshake<br/>QueryContextAttributes<br/>ValidateServerCert (opt)"]
    I["DoWebSocketHandshake<br/>HTTP upgrade + Sec-WebSocket-Accept validation"]
    J["Connected = True<br/>Stats reset"]

    A --> B --> C --> D --> E
    E -- YES --> F --> G
    E -- NO --> G
    G -- YES --> H --> I
    G -- NO --> I
    I --> J

    style A fill:#e1f5fe,stroke:#0277bd
    style J fill:#c8e6c9,stroke:#2e7d32
    style E fill:#fff9c4,stroke:#f9a825
    style G fill:#fff9c4,stroke:#f9a825
```

### Happy Eyeballs (RFC 6555)

The connection phase implements the Happy Eyeballs algorithm for dual-stack
hosts. When both IPv6 and IPv4 addresses are resolved:

1. IPv6 socket is created, set non-blocking, and `connect()` called immediately.
2. A 250ms race window starts. If IPv6 succeeds within this time, it wins.
3. If the race window expires, the IPv4 socket is also created.
4. Both sockets compete; the first to connect wins and the other is closed.
5. Fallback: if only one address family is available, it is used directly.

This guarantees the fastest possible connection while preferring IPv6 when
both are equally fast.

## TLS Handshake

The TLS layer is implemented through the Windows SSPI Schannel provider.
Wasabi performs the entire handshake manually rather than delegating to
WinHTTP or any higher-level abstraction.

![TLS Handshake](https://www.thesslstore.com/blog/wp-content/uploads/2017/01/SSL_Handshake_10-Steps-1.png)

### Credential Acquisition

Before the handshake begins, Wasabi initializes a `SCHANNEL_CRED` structure
with the following configuration:

| Field | Value | Purpose |
|:---|:---|:---|
| `dwVersion` | `SCHANNEL_CRED_VERSION` (4) | Structure version |
| `grbitEnabledProtocols` | `SP_PROT_TLS1_2_CLIENT \| SP_PROT_TLS1_3_CLIENT` | Accepted TLS versions |
| `dwFlags` | `SCH_CRED_NO_DEFAULT_CREDS` | Do not use Windows credential store |
| `dwFlags` | `SCH_CRED_MANUAL_CRED_VALIDATION` | Skip automatic certificate chain validation |
| `dwFlags` | `SCH_CRED_IGNORE_NO_REVOCATION_CHECK` | Do not fail if CRL is unavailable |
| `dwFlags` | `SCH_CRED_IGNORE_REVOCATION_OFFLINE` | Do not fail if CRL server is unreachable |

This credential is passed to `AcquireCredentialsHandle` with the package
name `"Microsoft Unified Security Protocol Provider"`.

> [!IMPORTANT]
> Certificate revocation checking is explicitly disabled (`IGNORE_NO_REVOCATION_CHECK` and `IGNORE_REVOCATION_OFFLINE`) to maximize compatibility with firewalled and offline corporate environments. This means that even if the certificate is issued by a trusted CA, the connection will proceed even if the CRL or OCSP responder is unreachable. Enabling strict revocation checking would require a registry change and is not recommended for client-side WebSocket connections in typical Office automation scenarios.

### Handshake Loop

The handshake is a multi-round exchange between the client and server.
The internal function `DoTLSHandshake` implements this as a loop.

```
Round 1: InitializeSecurityContext (first call, no input)
         → sends ClientHello
         → receives ServerHello + Certificate + ServerHelloDone

Round 2: InitializeSecurityContext (with server response)
         → sends ClientKeyExchange + ChangeCipherSpec + Finished
         → receives server ChangeCipherSpec + Finished

Result:  SEC_E_OK → handshake complete
```

Each round follows this pattern:

1. Call `InitializeSecurityContext` with the accumulated server data
2. If output token is produced, send it to the server via `sock_send`
3. If result is `SEC_I_CONTINUE_NEEDED`, read more data from the server
4. If result is `SEC_E_INCOMPLETE_MESSAGE`, read more data and retry
5. If any `SECBUFFER_EXTRA` is returned, preserve the extra bytes for the next round
6. If result is `SEC_E_OK`, the handshake is complete

The loop is protected by a maximum iteration count of 30 to prevent infinite
loops on malformed server responses.

### Post-Handshake

After the handshake completes, Wasabi queries the context for stream sizes
using `QueryContextAttributes` with `SECPKG_ATTR_STREAM_SIZES`. This returns:

| Field | Purpose |
|:---|:---|
| `cbHeader` | Size of the TLS record header (prepended to each encrypted block) |
| `cbTrailer` | Size of the TLS record trailer (appended to each encrypted block) |
| `cbMaximumMessage` | Maximum plaintext size per TLS record |

These values are used by `TLSSend` to correctly frame outgoing data.

## TLS Data Flow

### Encryption (Sending)

When `TLSSend` is called with plaintext data:

```mermaid
graph TD
    subgraph Plaintext["Plaintext Record"]
        P1["Header<br/>[cbHeader]"]:::header ---
        P2["Plaintext Data"]:::data ---
        P3["Trailer<br/>[cbTrailer]"]:::trailer
    end

    Plaintext --> EM["EncryptMessage"]:::process

    subgraph Encrypted["Encrypted Record"]
        E1["Header"]:::header ---
        E2["Encrypted Data"]:::data ---
        E3["Trailer"]:::trailer
    end

    EM --> Encrypted

    Encrypted --> SS["sock_send"]:::send

    classDef header fill:#fff9c4,stroke:#f9a825,color:#333
    classDef data fill:#e1f5fe,stroke:#0277bd,color:#333
    classDef trailer fill:#f3e5f5,stroke:#7b1fa2,color:#333
    classDef process fill:#c8e6c9,stroke:#2e7d32,color:#333
    classDef send fill:#ffcc80,stroke:#e65100,color:#333
    style Plaintext fill:none,stroke:#333,stroke-width:1px
    style Encrypted fill:none,stroke:#333,stroke-width:1px
```

`TLSSend` automatically splits data larger than `cbMaximumMessage` into multiple TLS records, each encrypted separately and sent sequentially.

### Decryption (Receiving)

When `TLSDecrypt` processes buffered data:

```mermaid
graph TD
    A["RecvBuffer<br/>(raw bytes from socket)"]
    B["DecryptMessage"]
    C["SECBUFFER_DATA<br/>(decrypted)"]
    D["SECBUFFER_EXTRA<br/>(next record)"]
    E["Append to<br/>DecryptBuffer"]
    F["Move to start<br/>of RecvBuffer"]

    A --> B
    B --> C
    B --> D
    C --> E
    D --> F

    style A fill:#e1f5fe,stroke:#0277bd
    style B fill:#fff9c4,stroke:#f9a825
    style C fill:#c8e6c9,stroke:#2e7d32
    style D fill:#f3e5f5,stroke:#7b1fa2
    style E fill:#c8e6c9,stroke:#2e7d32
    style F fill:#f3e5f5,stroke:#7b1fa2
```

The `SECBUFFER_EXTRA` handling is critical. When the OS socket delivers
multiple TLS records in a single `recv()` call, `DecryptMessage` only
processes the first complete record and flags the remaining bytes as
`SECBUFFER_EXTRA`. Wasabi moves these bytes to the beginning of
`RecvBuffer` and loops to decrypt again.

> [!NOTE]
> If `DecryptMessage` returns `SEC_E_INCOMPLETE_MESSAGE`, the current
> `RecvBuffer` does not contain a complete TLS record. Wasabi exits the
> decrypt loop and waits for more data to arrive on the next polling cycle.

## WebSocket Handshake

After the TCP connection (and optional TLS handshake) is established, Wasabi
performs the WebSocket protocol upgrade as defined by
[RFC 6455 Section 4](https://datatracker.ietf.org/doc/html/rfc6455#section-4).

![WebSocket Handshake](https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRuz8ayWEYgqlNTWFMwLvocD7CsCBpVr3iP57zVbCGyuKHtpnKlBS_chhU&s=10)

### Request

Wasabi constructs and sends an HTTP/1.1 GET request:

```http
GET /path HTTP/1.1
Host: example.com
Upgrade: websocket
Connection: Upgrade
Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==
Sec-WebSocket-Version: 13
Origin: https://example.com
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36
```

The `Sec-WebSocket-Key` is a Base64-encoded 16-byte random value generated
by `GenerateWSKey`. Wasabi uses `CryptGenRandom` (from `advapi32.dll`) to
obtain cryptographically strong randomness for this key. If `CryptGenRandom`
were to fail (which should never happen on a normal system), the code falls
back to the VBA `Rnd` function.

If custom headers, subprotocol, or proxy credentials are configured, they
are appended before the final blank line.

### Response Validation

Wasabi validates the server response in two steps:

**1. Status code check**

The response must contain HTTP status `101` (Switching Protocols). Any
other status triggers `ERR_HANDSHAKE_REJECTED`.

**2. Accept key validation**

Wasabi computes the expected accept value:

```
expected = Base64(SHA1(key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"))
```

This is compared against the `Sec-WebSocket-Accept` header in the server
response. A mismatch triggers `ERR_HANDSHAKE_REJECTED`.

The SHA-1 implementation is internal to Wasabi and does not depend on any
external library. See the [SHA-1 section](#sha-1-implementation) below.

## Frame Processing

WebSocket communication happens through frames as defined by
[RFC 6455 Section 5](https://datatracker.ietf.org/doc/html/rfc6455#section-5).

### Frame Format

```mermaid
graph LR
    subgraph Frame["WebSocket Frame (RFC 6455)"]
        direction LR
        B0["FIN (1) RSV (3) Opcode (4)"]
        B1["MASK (1) Payload len (7)"]
        Ext16["Extended len (16)"]
        Ext64["Extended len (64)"]
        Mask["Masking Key (32)"]
        Payload["Payload Data"]
    end

    B0 --> B1
    B1 -->|len=126| Ext16
    B1 -->|len=127| Ext64
    B1 -->|len<126| Payload
    Ext16 --> Payload
    Ext64 --> Payload
    B1 -->|MASK=1| Mask
    Mask --> Payload

    style Frame fill:none,stroke:#333,stroke-width:2px
    style B0 fill:#e1f5fe,stroke:#0277bd
    style B1 fill:#e1f5fe,stroke:#0277bd
    style Ext16 fill:#fff9c4,stroke:#f9a825
    style Ext64 fill:#fff9c4,stroke:#f9a825
    style Mask fill:#f3e5f5,stroke:#7b1fa2
    style Payload fill:#c8e6c9,stroke:#2e7d32
```

### Outgoing Frames (Sending)

When `WebSocketSend` or `WebSocketSendBinary` is called:

1. The payload is measured in bytes (UTF-8 for text, raw for binary)
2. A 4-byte cryptographically random mask key is generated via `CryptGenRandom` (`FillRandomBytes`)
3. The frame header is constructed with the FIN bit set, the appropriate
   opcode (`0x01` for text, `0x02` for binary), and the MASK bit set
4. The payload length is encoded in the appropriate tier
5. Each payload byte is XORed with `mask(i Mod 4)`
6. The complete frame is sent via `RawSendFor` (or `TLSSend` for TLS)

### Incoming Frames (Receiving)

The internal function `ProcessFrames` parses frames from the
`DecryptBuffer`:

```mermaid
graph TD
    A[DecryptBuffer]
    B["Read byte 0<br/>FIN bit + opcode"]
    C["Read byte 1<br/>MASK bit + payload length"]
    D{"Payload length?"}
    E["Use as-is<br/>length &lt; 126"]
    F["Read 2 more bytes<br/>length = 126"]
    G["Read 8 more bytes<br/>length = 127"]
    H{"MASK bit set?"}
    I["Read 4-byte mask key"]
    J["Extract payload<br/>using CopyMemory"]
    K{"Opcode?"}
    L["0x00 Continuation<br/>append to fragment buffer"]
    M["0x01 Text<br/>UTF-8 decode, enqueue"]
    N["0x02 Binary<br/>enqueue raw bytes"]
    O["0x08 Close<br/>send close response, disconnect"]
    P["0x09 Ping<br/>send pong with same payload"]

    A --> B --> C --> D
    D -- "&lt;126" --> E --> H
    D -- "126" --> F --> H
    D -- "127" --> G --> H
    H -- YES --> I --> J
    H -- NO --> J
    J --> K
    K -- "0x00" --> L
    K -- "0x01" --> M
    K -- "0x02" --> N
    K -- "0x08" --> O
    K -- "0x09" --> P

    style A fill:#e1f5fe,stroke:#0277bd
    style D fill:#fff9c4,stroke:#f9a825
    style H fill:#fff9c4,stroke:#f9a825
    style K fill:#fff9c4,stroke:#f9a825
    style L fill:#f3e5f5,stroke:#7b1fa2
    style M fill:#c8e6c9,stroke:#2e7d32
    style N fill:#c8e6c9,stroke:#2e7d32
    style O fill:#ffcdd2,stroke:#c62828
    style P fill:#ffe0b2,stroke:#e65100
```

### Fragmentation

Large messages may arrive split across multiple frames. The first frame
has a non-zero opcode and the FIN bit cleared. Continuation frames use
opcode `0x00`. The final frame has the FIN bit set.

```
Frame 1: FIN=0, opcode=0x01 (Text), payload="Hello "
Frame 2: FIN=0, opcode=0x00 (Continuation), payload="from "
Frame 3: FIN=1, opcode=0x00 (Continuation), payload="Wasabi"

Result: "Hello from Wasabi"
```

Wasabi accumulates fragments in the per-connection `FragmentBuffer` using
`CopyMemory`. When the final FIN frame arrives, the complete payload is
assembled and enqueued as a single message.

> [!NOTE]
> The fragment buffer size defaults to 256KB and can be configured via
> `WebSocketSetBufferSizes` before connecting.

## Message Queues

Each connection maintains two independent circular queues (ring buffers):
one for text messages and one for binary messages.

![Ring Buffer](https://www.intel.com/content/dam/developer/articles/technical/fast-core-to-core-communications/ring-buffer-arch.png)

### Structure

```mermaid
graph LR
    subgraph Queue["Ring Buffer (512 entries)"]
        direction LR
        S0[0]:::empty --- S1[1]:::empty --- S2["M1"]:::filled --- S3["M2"]:::filled --- S4["M3"]:::filled --- S5["M4"]:::filled --- S6[6]:::empty --- S7[7]:::empty
    end

    H["Head"]:::pointer --> S2
    T["Tail"]:::pointer --> S6

    Info["MsgCount = 4"]

    classDef empty fill:#bdc3c7,stroke:#95a5a6,color:#333
    classDef filled fill:#c8e6c9,stroke:#2e7d32,color:#333
    classDef pointer fill:#fff9c4,stroke:#f9a825,color:#333
    style Queue fill:none,stroke:#3498db,stroke-width:2px
    style Info fill:none,stroke:none,color:#333
```

### Operations

**Enqueue (new message arrives):**
```
MsgQueue(MsgTail) = message
MsgTail = (MsgTail + 1) Mod MSG_QUEUE_SIZE
MsgCount = MsgCount + 1
```

**Dequeue (WebSocketReceive called):**
```
result = MsgQueue(MsgHead)
MsgHead = (MsgHead + 1) Mod MSG_QUEUE_SIZE
MsgCount = MsgCount - 1
```

Both operations are O(1) with no memory allocation or copying beyond the
initial array setup.

> [!WARNING]
> When `MsgCount` reaches `MSG_QUEUE_SIZE` (512), new messages are
> discarded and a warning is logged.

## Receive Pipeline

The complete data flow from network to your code.

```mermaid
graph TD
    A[Network]
    B["sock_recv"]
    C["tempBuf"]
    D{"Connection type?"}
    E["RecvBuffer<br/>↓↓<br/>DecryptMessage<br/>↓↓<br/>DecryptBuffer"]
    F["DecryptBuffer<br/>(directly)"]
    G["ProcessFrames"]
    H{"Frame opcode?"}
    I["Text frame<br/>Utf8ToString<br/>MsgQueue enqueue"]
    J["Binary frame<br/>BinaryQueue enqueue"]
    K["Ping<br/>SendPongFrame<br/>(automatic response)"]
    L["Close<br/>WebSocketSendClose<br/>+ disconnect"]
    M["Continuation<br/>FragmentBuffer<br/>accumulation"]
    N["WebSocketReceive"]
    O["Your VBA code"]

    A --> B --> C --> D
    D -- "TLS" --> E --> G
    D -- "Plain" --> F --> G
    G --> H
    H -- "Text" --> I
    H -- "Binary" --> J
    H -- "Ping" --> K
    H -- "Close" --> L
    H -- "Cont." --> M
    I --> N
    J --> N
    K --> N
    L --> N
    M --> N
    N --> O

    style A fill:#e1f5fe,stroke:#0277bd
    style D fill:#fff9c4,stroke:#f9a825
    style H fill:#fff9c4,stroke:#f9a825
    style I fill:#c8e6c9,stroke:#2e7d32
    style J fill:#c8e6c9,stroke:#2e7d32
    style K fill:#ffe0b2,stroke:#e65100
    style L fill:#ffcdd2,stroke:#c62828
    style M fill:#f3e5f5,stroke:#7b1fa2
    style N fill:#e1f5fe,stroke:#0277bd
    style O fill:#e1f5fe,stroke:#0277bd
```

### Buffer Sizes

| Buffer | Default Size | Configurable |
|:---|:---|:---|
| `RecvBuffer` | 256KB | Yes, via `WebSocketSetBufferSizes` |
| `DecryptBuffer` | 256KB | Yes, via `WebSocketSetBufferSizes` |
| `FragmentBuffer` | 256KB | Yes, via `WebSocketSetBufferSizes` |
| Text queue | 512 entries | No (compile-time constant) |
| Binary queue | 512 entries | No (compile-time constant) |

## Auto-Reconnect

When a connection loss is detected during polling and auto-reconnect is
enabled, Wasabi executes the following sequence.

```mermaid
graph TD
    A["Connection lost detected"]
    B["Save all settings<br/>(URL, proxy, headers, subprotocol,<br/>timeouts, callbacks, NoDelay)"]
    C["CleanupHandle<br/>(close socket, release TLS, clear buffers)"]
    D["Calculate delay<br/>delay = baseDelay * 2^(attempt-1)<br/>cap at MAX_RECONNECT_DELAY_MS (30s)"]
    E["Wait<br/>(DoEvents loop)"]
    F["Reallocate buffers"]
    G["Restore all saved settings"]
    H["ConnectHandle(handle, savedUrl)"]
    I["Success → ReconnectAttempts = 0"]
    J["Failure → increment attempt counter<br/>try again if under max"]

    A --> B --> C --> D --> E --> F --> G --> H
    H --> I
    H --> J

    style A fill:#ffcdd2,stroke:#c62828
    style B fill:#e1f5fe,stroke:#0277bd
    style C fill:#e1f5fe,stroke:#0277bd
    style D fill:#fff9c4,stroke:#f9a825
    style E fill:#fff9c4,stroke:#f9a825
    style F fill:#e1f5fe,stroke:#0277bd
    style G fill:#e1f5fe,stroke:#0277bd
    style H fill:#fff9c4,stroke:#f9a825
    style I fill:#c8e6c9,stroke:#2e7d32
    style J fill:#ffcdd2,stroke:#c62828
```

### Backoff Pattern

| Attempt | Delay (base = 1000ms) |
|:---|:---|
| 1 | 1000ms |
| 2 | 2000ms |
| 3 | 4000ms |
| 4 | 8000ms |
| 5 | 16000ms |
| 6+ | 30000ms (capped) |

> [!IMPORTANT]
> The reconnect delay loop uses `DoEvents`, which yields to the Windows
> message pump but does not fully release the VBA thread. The Office UI
> remains partially responsive during this wait.

## SHA-1 Implementation

Wasabi includes a complete SHA-1 implementation in pure VBA for computing
the `Sec-WebSocket-Accept` header during the WebSocket handshake, as
required by [RFC 6455 Section 4.2.2](https://datatracker.ietf.org/doc/html/rfc6455#section-4.2.2).

### Why Internal

The SHA-1 hash is needed exactly once per connection. Using an external
dependency (such as `ScriptControl` or a COM hash object) would break
Wasabi's zero-dependency constraint. The internal implementation ensures
the module remains a single portable `.bas` file.

### Unsigned 32-bit Arithmetic

VBA's `Long` type is a signed 32-bit integer. SHA-1 requires unsigned
32-bit addition and rotation. Wasabi works around this with three helper
functions:

| Function | Purpose |
|:---|:---|
| `ADD32(a, b)` | Unsigned 32-bit addition using split high/low halves |
| `ROTL32(v, n)` | Left rotation by n bits using iterative shift and carry |
| `U32Shl1(v)` | Single-bit left shift handling the sign bit explicitly |

These functions use bitmasks (`&H7FFF`, `&HFFFF`, `&H80000000`) to
isolate and manipulate individual bit ranges without triggering VBA
overflow errors.

### SHA-1 Constants

| Round Range | Constant | Hex |
|:---|:---|:---|
| 0 to 19 | `0x5A827999` | `&H5A827999` |
| 20 to 39 | `0x6ED9EBA1` | `&H6ED9EBA1` |
| 40 to 59 | `0x8F1BBCDC` | `&H8F1BBCDC` |
| 60 to 79 | `0xCA62C1D6` | `&HCA62C1D6` |

### Initial Hash Values

```
h0 = 0x67452301
h1 = 0xEFCDAB89
h2 = 0x98BADCFE
h3 = 0x10325476
h4 = 0xC3D2E1F0
```

## Proxy Tunnel

When proxy is enabled, Wasabi establishes an HTTP CONNECT tunnel (or a
SOCKS5 tunnel) before performing TLS or WebSocket handshaking.

```mermaid
sequenceDiagram
    participant C as Client
    participant P as Proxy
    participant S as Server

    C->>+P: CONNECT host:port HTTP/1.1<br/>Host: host:port<br/>[Proxy-Authorization]
    P->>+S: TCP connect
    P-->>-C: HTTP/1.1 200 Connection Established

    Note over C,S: TLS Handshake through tunnel
    Note over C,S: WebSocket Handshake through tunnel
    Note over C,S: WebSocket Frames through tunnel
```

> [!NOTE]
> The proxy only sees the CONNECT request. After the tunnel is established,
> all subsequent traffic (TLS, WebSocket) passes through opaquely. The proxy
> cannot inspect the encrypted content.

## Maintenance Cycle

Every call to `WebSocketReceive` triggers an internal maintenance pass
via `TickMaintenance`. This is the only mechanism for time-based
features because VBA does not support background timers.

### What Maintenance Checks

| Check | Condition | Action |
|:---|:---|:---|
| Automatic Ping | `PingIntervalMs > 0` and interval elapsed | Send Ping frame |
| Inactivity Timeout | `InactivityTimeoutMs > 0` and threshold exceeded | Close connection, trigger reconnect if enabled |
| MTU Probe | `AutoMTU` and `ProbeEnabled` and interval elapsed | Call `ProbeMTU` to re-measure MSS |

> [!IMPORTANT]
> If your code stops calling `WebSocketReceive`, maintenance also stops.
> Automatic pings will not be sent and inactivity timeouts will not fire.

## Memory Layout

Wasabi uses pre-allocated byte arrays instead of dynamic string
concatenation to minimize heap fragmentation in long-running sessions.

```
Per connection memory footprint (default settings):

  RecvBuffer:      256 KB
  DecryptBuffer:   256 KB
  FragmentBuffer:  256 KB
  MsgQueue:        512 × String pointer
  BinaryQueue:     512 × Byte array pointer
  CustomHeaders:   32 × String pointer

  Total baseline: ~768 KB + queue overhead per connection
  Maximum (64 connections): ~48 MB
```

> [!NOTE]
> The actual memory consumed depends on the size of queued messages and
> the configured buffer sizes. The baseline above represents the fixed
> allocation before any messages are received.

## Error Propagation

Errors in Wasabi propagate through two parallel paths.

```mermaid
graph TD
    A[Internal error occurs]
    B["Per-connection state<br/>m_Connections(h).LastError<br/>m_Connections(h).LastErrorCode<br/>m_Connections(h).TechnicalDetails"]
    C["Global state<br/>m_LastError<br/>m_LastErrorCode<br/>m_TechnicalDetails"]
    D["WebSocketGetLastError(h)<br/>(handle-specific)"]
    E["WebSocketGetLastError()<br/>(global fallback)"]

    A --> B
    A --> C
    B --> D
    C --> E

    style A fill:#ffcdd2,stroke:#c62828
    style B fill:#e1f5fe,stroke:#0277bd
    style C fill:#e1f5fe,stroke:#0277bd
    style D fill:#c8e6c9,stroke:#2e7d32
    style E fill:#fff9c4,stroke:#f9a825
```

When a function is called with a valid handle, the per-connection error
state is returned. When called without a handle or with an invalid handle,
the global error state is returned.

> [!TIP]
> Always pass the handle when checking errors to get the most specific
> information.

## Related Documentation

- [API Reference](API_REFERENCE.md) for the complete public API
- [Error Reference](ERRORS.md) for detailed error diagnostics
- [SECURITY.md](../SECURITY.md) for security design decisions
- [RFC 6455](https://datatracker.ietf.org/doc/html/rfc6455) for the WebSocket protocol specification
- [SSPI/Schannel documentation](https://learn.microsoft.com/en-us/windows/win32/secauthn/sspi) for the TLS implementation reference
- [Winsock documentation](https://learn.microsoft.com/en-us/windows/win32/winsock/winsock-functions) for the transport layer reference
