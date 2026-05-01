<div align="center">
  <img src="resources/logo.png" width="150" />
</div>

<h1 align="center">Wasabi</h1>

<p align="center">
  Production-ready WebSocket and WSS for VBA with native TLS, auto reconnect, proxy support, and zero external dependencies
</p>

<p align="center">
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/architecture-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/TLS-1.2%20%2F%201.3-brightgreen.svg" alt="TLS" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/WebSocket-RFC%206455-orange.svg" alt="WebSocket" />
  <img src="https://img.shields.io/badge/Proxy-HTTP%20%2B%20SOCKS5-yellowgreen" alt="Proxy" />
  <img src="https://img.shields.io/badge/mTLS-PFX%20%2B%20Store-yellow" alt="mTLS" />
  <img src="https://img.shields.io/badge/MQTT-3.1.1%20over%20WS-purple" alt="MQTT" />
  <img src="https://img.shields.io/badge/Proxy%20Auth-NTLM%2FKerberos-red" alt="NTLM" />
  <img src="https://img.shields.io/badge/revocation-CRL%2FOCSP-lightgrey" alt="Revocation" />
  <img src="https://img.shields.io/github/stars/uesleibros/wasabi?style=flat&color=gold" alt="Stars" />
  <img src="https://img.shields.io/github/last-commit/uesleibros/wasabi?style=flat" alt="Last Commit" />
</p>

## What is Wasabi

Wasabi is a VBA module designed to make WebSocket communication simple, predictable, and practical bringing an experience similar to [socket.io](https://socket.io) in Node.js, but entirely within the Office ecosystem. It is a single, self-contained `.bas` file that compiles seamlessly on 32-bit and 64-bit Office hosts, from Windows XP to Windows 11.

## Roadmap

- [x] IPv6 and SNI support
- [x] Mutual TLS (mTLS) for client certificate authentication
- [x] SOCKS5 proxy support
- [ ] WebSocket over HTTP/2
- [ ] NTLM/Kerberos for proxies
- [ ] MQTT over WebSocket
- [ ] RTT measurement (GetLatency)
- [ ] permessage-deflate compression (RFC 7692)
- [ ] I/O Completion Ports (IOCP) for kernel-driven socket monitoring
- [x] Zero-copy receive buffers
- [x] MTU-aware frame sizing
- [x] Send batching
- [x] Close frame payload parsing
- [x] Happy Eyeballs (RFC 6555)
- [x] CRL/OCSP certificate revocation checking

## Why Wasabi exists

VBA is excellent for automation and integration with Excel, PowerPoint, Word and other Office applications, but it hits a wall when real-time communication is required. In practice, anyone trying to build a project that depends on live messaging usually runs into three problems:

- **No standardization:** there is no official, modern path for sockets and WebSockets in VBA.
- **Verbose low-level APIs:** the most common options require a lot of infrastructure code just to connect and maintain a stable session.
- **Limited event-driven patterns:** it is common to end up with loops, timers and control logic just to simulate something that other languages handle natively.

## What VBA limitations it solves

Working with networking in VBA often becomes a project of its own. Some typical pain points:

**Winsock and Windows API calls**
- Require declarations, structs, callbacks and low-level details unrelated to the actual goal.
- Small adjustments can break compatibility or introduce hard-to-track bugs.

**Security and randomness**
- WebSocket connections require cryptographically secure XOR masks to protect against proxy cache poisoning attacks.
- VBA's built-in `Rnd` function is deterministic and unsuitable for this purpose.
- Wasabi uses `CryptGenRandom` from `advapi32.dll` for all WebSocket frame masks, falling back to `Rnd` only in the extraordinary case that the cryptographic API is unavailable.

**HTTP is not WebSocket**
- Even with WinHTTP or MSXML, you are in a request/response world.
- Real-time scenarios turn into polling, long-polling or workarounds that consume resources and increase latency.

**Limited asynchronism**
- VBA was not designed for modern concurrency.
- Without a good abstraction, it is easy to freeze the UI or create inconsistent state.

**Maintenance and readability**
- Most handcrafted solutions grow long and fragile.
- The networking layer becomes the largest part of the project, making simple things hard to maintain.

## Where it is useful

- **Bots** (Discord, Slack, Telegram), connect to gateways and handle real-time events and messages directly from Excel or Word
- **Trading and finance**, stream live prices from exchanges like Binance, Coinbase or B3 into spreadsheet cells with millisecond-level latency
- **Dashboards**, update live data on a spreadsheet without manual refresh or polling HTTP endpoints
- **IoT and industrial**, receive sensor data from ESP32, Raspberry Pi or SCADA systems via WebSocket directly into Office
- **Games and interactive tools**, build client/server communication for VBA-based games or collaborative tools
- **Corporate automation**, connect Office to internal WebSocket APIs behind proxies and firewalls without installing anything

## Quick Start

> [!NOTE]
> Before using Wasabi, it is highly recommended that you review the [documentation](docs).

### Import

[Download the latest version of Wasabi](../../releases) and import it into your VBA project via
**File → Import File** in the VBA editor.

No references need to be enabled in **Tools → References**.

> For the complete reference with examples, parameters, return values, and usage
notes, see [API Reference](docs/API_REFERENCE.md).

## Compatibility

Wasabi was designed to run without any external dependencies, using exclusively
native Windows DLLs that ship with every version of Windows. No references need
to be enabled in **Tools → References**, no COM components need to be registered,
and no third-party installers are required. Dropping the `.bas` file into a VBA
project is all it takes.

### Operating System

| Version | Support |
|---|---|
| Windows XP | ✅ |
| Windows Vista | ✅ |
| Windows 7 | ✅ |
| Windows 8 / 8.1 | ✅ |
| Windows 10 | ✅ |
| Windows 11 | ✅ |

Wasabi relies exclusively on `ws2_32.dll` (Winsock 2), `secur32.dll` (Schannel
SSPI), `kernel32.dll` and `advapi32.dll`. These libraries have been present and
stable in every version of Windows since XP, which is why Wasabi can run on
machines that are over 20 years old without any modifications.

This is a deliberate architectural choice. Many competing modules depend on the
`WinHttpWebSocket*` family of functions (`WinHttpWebSocketSend`,
`WinHttpWebSocketReceive`, `WinHttpWebSocketCompleteUpgrade`) introduced only in
Windows 8. As a result, those modules silently fail on Windows 7 machines, which
remain common in corporate and industrial environments. Wasabi has no such
limitation.

### Office and VBA

| Environment | Support |
|---|---|
| Excel 32-bit | ✅ |
| Excel 64-bit | ✅ |
| Word 32-bit | ✅ |
| Word 64-bit | ✅ |
| PowerPoint 32-bit | ✅ |
| PowerPoint 64-bit | ✅ |
| Access 32-bit | ✅ |
| Access 64-bit | ✅ |
| Any VBA7 host (Office 2010+) | ✅ |
| VBA6 (Office 2007 and earlier) | ✅ |

The transition from 32-bit to 64-bit Office broke a large number of VBA modules
that used native API declarations, because pointer sizes changed from 4 bytes to
8 bytes. Wasabi handles this transparently through conditional compilation.

Every single API declaration in the module uses the `#If VBA7` compiler
directive to switch between two complete sets of declarations: one using `Long`
for 32-bit environments and one using `LongPtr` and `PtrSafe` for 64-bit
environments. This means the same unmodified `.bas` file works correctly whether
the user is running Office 2007 on a 32-bit machine or Office 365 64-bit on
Windows 11.

> 32-bit and 64-bit compatibility is guaranteed through conditional compilation (`#If VBA7`) across all API declarations. `LongPtr` and `PtrSafe` are applied automatically at compile time.

### Native DLLs

| Library | Role in Wasabi |
|---|---|
| `ws2_32.dll` | TCP socket creation, DNS resolution, connection, send, recv, I/O control |
| `secur32.dll` | TLS 1.2 and TLS 1.3 via Schannel SSPI (handshake, encryption, decryption) |
| `kernel32.dll` | Memory operations, UTF-8 string conversion, tick count for timeouts |
| `advapi32.dll` | Cryptographic random number generation (`CryptGenRandom`) for secure WebSocket frame masking |

**ws2_32.dll (Windows Sockets 2)**
This is the core networking library of Windows. Wasabi uses it directly to
create TCP sockets, resolve hostnames via `gethostbyname`, establish connections,
send and receive raw bytes, and control socket behavior through `ioctlsocket` and
`setsockopt`. By going to this layer directly, Wasabi avoids the performance and
flexibility limitations of higher-level abstractions like WinHTTP.

**secur32.dll (Security Support Provider Interface)**
This is the Windows security library responsible for TLS. Wasabi uses it to
perform the full TLS handshake manually via `AcquireCredentialsHandle` and
`InitializeSecurityContext`, and to encrypt and decrypt data after the handshake
via `EncryptMessage` and `DecryptMessage`. This gives Wasabi complete control
over TLS flags, protocol versions, certificate validation behavior and cipher
negotiation, something that is impossible when delegating to WinHTTP.

**kernel32.dll**
Used for three specific purposes: `RtlMoveMemory` (exposed as `CopyMemory`) for
direct buffer manipulation without VBA overhead, `MultiByteToWideChar` and
`WideCharToMultiByte` for correct UTF-8 encoding and decoding of WebSocket text
frames, and `GetTickCount` for measuring elapsed time in timeout and reconnect
logic, with wraparound handling for systems with long uptimes.

**advapi32.dll**
Used exclusively for `CryptGenRandom`, which generates cryptographically secure
random bytes. Wasabi uses this to create the 4-byte XOR masks required by the
WebSocket protocol for every outgoing frame. This is a critical security measure:
predictable masks can be exploited by malicious intermediaries to poison caching
proxies or inject data into the connection. If `CryptGenRandom` is unavailable
(which should never happen on a normal Windows installation), Wasabi falls back
to VBA's `Rnd` function with a suitable warning.

### What does "zero external dependencies" mean in practice

In practical terms, Wasabi does not require anything beyond Windows and the VBA
runtime you already have.

There is no installer, no registered COM component, no ActiveX control, no
third-party DLL, no Python runtime, no .NET package, no `regsvr32`, and no
extra setup step outside importing the `.bas` file into your project.

This matters a lot in corporate environments, where IT policies often prevent
users from installing software, registering components, or adding external
libraries to Office solutions. Many networking alternatives for VBA depend on
COM objects, commercial SDKs, or extra binaries that require administrator
rights, licensing, or machine-specific configuration.

Wasabi avoids all of that by relying only on native Windows libraries that are
already present on the machine, such as `ws2_32.dll`, `secur32.dll`, and
`kernel32.dll`. If Windows is running, the underlying networking and TLS stack
required by Wasabi is already there.

The result is simple: you can distribute a workbook or VBA project containing
Wasabi without asking the user to install anything first. No setup wizard, no
dependency hell, no missing runtime, no registration step, and no surprise
failure because a component was not deployed correctly.

## Execution Model: Single-Thread and Polling

VBA is a single-threaded language. There is exactly one execution thread, shared
between your code and the Office interface. This means there is no native way to
listen to a socket in the background while the user interacts.

Wasabi solves this with a polling model: instead of dispatching events
automatically when a message arrives, it stores incoming messages in an internal
Ring Buffer and waits for your code to retrieve them.

Each time you call `WebSocketReceive`, Wasabi does three things:

1. Checks whether data is available on the socket (`FIONREAD`)
2. Reads and processes any frames received since the last call
3. Returns the next message from the queue, or an empty string if there is nothing

Between calls, the socket remains fully open. The Windows kernel continues
buffering incoming data at the driver level regardless of what your VBA code is
doing. No messages are lost between calls.

### What this does not mean

- **Not slow.** The Ring Buffer holds up to 512 messages. Nothing is lost between
calls.
- **Connection does not drop.** The socket stays open and active regardless of
polling frequency.
- **No infinite loop required.** One-shot use cases (connect, send, receive,
disconnect) work fine with a simple wait loop.

## Acknowledgements

Wasabi did not emerge from nothing. It was built on years of attempts, workarounds
and creative solutions that the VBA community constructed over time to solve a
problem the language never officially addressed.

The following projects were essential, not only as technical references but as
proof that serious network communication inside the Office ecosystem was possible.
Each of them, in their own way, pushed the boundary of what was considered viable
in VBA.

- [**WinHttpWebSocket**](https://github.com/EagleAglow/vba-websocket), one of the first serious attempts at WebSocket in VBA using native Windows APIs,
  and the most widely referenced starting point in the community.

- [**VbAsyncSocket**](https://github.com/wqweto/VbAsyncSocket), the most technically sophisticated VB6/VBA networking library available,
  with native Schannel support and a level of engineering depth that set the bar
  for what was possible in the ecosystem.

- [**VBA-Web**](https://github.com/VBA-tools/VBA-Web), standard for HTTP and REST communication in VBA, and a reference
  for how a well-documented, well-maintained VBA library should be structured.

- [**TlsSocketWSS**](https://github.com/Maatooh/TlsSocketWSS-vb6), simple TlsSocket server implemented to work with websocket, allowing connection from WSS with a self-signed certificate or a valid one.

- [**VB6-WebSocket-Server-SSL**](https://github.com/JoshyFrancis/vb6-websocket-server-ssl), secure WebSocket in Pure VB6 no external Libraries or ActiveX.

- [**VBA_WinSockAPI**](https://github.com/papanda925/VBA_WinsockAPI_TCP_Sample), a clean, educational TCP client/server sample using raw Winsock API calls
  directly from VBA, demonstrating the fundamentals of socket creation, binding,
  listening, accepting and sending/receiving data that many VBA networking
  projects build upon.
  
Studying what these projects did well, and especially what they could not solve,
shaped every architectural decision in Wasabi: the non-blocking I/O model, the
manual Schannel implementation, the Ring Buffers, the Auto-Reconnect. None of it
would have taken form without the foundation these developers built first.

The VBA community is smaller than it deserves to be, and anyone who chooses to
spend time on an open project in this ecosystem is doing something genuinely
valuable. Wasabi exists on top of that work, and that will never be a small thing.

## Contributing

Bug reports, feature requests, and pull requests are welcome. Please read
[CONTRIBUTING.md](CONTRIBUTING.md) before contributing.

## Security

Please do **not** report security vulnerabilities through public issues. See
[SECURITY.md](SECURITY.md) and use GitHub Private Vulnerability Reporting.

## License

**MIT**, free for personal and commercial use. See [LICENSE](LICENSE) for details.
