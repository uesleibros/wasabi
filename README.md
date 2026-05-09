<div align="center">
  <img src="resources/logo.png" width="150" />
</div>

<h1 align="center">Wasabi - VBA WebSocket & TCP</h1>

<p align="center">
  <b>Real-time WebSocket, MQTT, and raw TCP for Microsoft Office. No dependencies, no COM, no installs.</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <a href="dev/asm">
    <img src="https://img.shields.io/badge/Engine-Wasabi%20ASM-2ecc71?style=flat&logo=webassembly&logoColor=lightgreen" alt="Engine" />
  </a>
  <img src="https://img.shields.io/badge/Performance-Low--Level-orange?style=flat&logo=speedtest&logoColor=white" alt="Performance" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/architecture-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/TLS-1.2%20%2F%201.3-brightgreen.svg" alt="TLS" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/WebSocket-RFC%206455-orange.svg" alt="WebSocket" />
  <img src="https://img.shields.io/badge/Proxy-Auto--Discovery-yellowgreen" alt="Proxy" />
  <img src="https://img.shields.io/badge/mTLS-PFX%20%2B%20Store-yellow" alt="mTLS" />
  <img src="https://img.shields.io/badge/MQTT-5%20%26%20QoS%201%2F2-purple" alt="MQTT" />
  <img src="https://img.shields.io/badge/Resilience-Offline%20Queue-success" alt="Offline Queue" />
  <img src="https://img.shields.io/badge/Proxy%20Auth-NTLM%2FKerberos-red" alt="NTLM" />
  <img src="https://img.shields.io/badge/RTT-latency%20measurement-orange" alt="RTT" />
  <img src="https://img.shields.io/badge/Deflate-permessage--deflate-success" alt="Deflate" />
  <img src="https://img.shields.io/badge/TCP-Native%20Client-blue" alt="TCP" />
  <img src="https://img.shields.io/badge/Middleware-Pipeline-blueviolet" alt="Middleware" />
  <img src="https://img.shields.io/badge/Compression-Pluggable-red" alt="Compression Pluggable" />
  <img src="https://img.shields.io/badge/Async-Event--Driven-teal" alt="Async" />
  <img src="https://img.shields.io/github/stars/uesleibros/wasabi?style=flat&color=gold" alt="Stars" />
  <img src="https://img.shields.io/github/last-commit/uesleibros/wasabi?style=flat" alt="Last Commit" />
  <a href="../../releases">
    <img src="https://img.shields.io/github/downloads/uesleibros/wasabi/total.svg?style=flat" alt="Downloads" />
  </a>
  <a href="../../releases">
    <img src="https://img.shields.io/github/v/release/uesleibros/wasabi?style=flat" alt="Latest Version" />
  </a>
</p>

<h2 align="center">Trusted by the VBA Community</h2>

<p align="center">
  <a href="https://github.com/EagleAglow">
    <img src="https://github.com/EagleAglow.png?size=90" width="90" alt="EagleAglow" />
  </a>
  <a href="https://github.com/wqweto">
    <img src="https://github.com/wqweto.png?size=90" width="90" alt="wqweto" />
  </a>
  <a href="https://github.com/Maatooh">
    <img src="https://github.com/Maatooh.png?size=90" width="90" alt="Maatooh" />
  </a>
  <a href="https://github.com/PerditionC">
    <img src="https://github.com/PerditionC.png?size=90" width="90" alt="PerditionC" />
  </a>
  <a href="https://github.com/Sven-Bo">
    <img src="https://github.com/Sven-Bo.png?size=90" width="90" alt="Sven-Bo" />
  </a>
  <a href="https://github.com/DecimalTurn">
    <img src="https://github.com/DecimalTurn.png?size=90" width="90" alt="DecimalTurn" />
  </a>
</p>

<p align="center">
  <sub>
    Recognition from developers behind <br/> 
    influential VBA/VB6 systems, and/or low-level infrastructure projects.
  </sub>
</p>

> [!NOTE]
> **Supported Applications**
>
> ![](resources/svg/ms-powerpoint.svg)
> ![](resources/svg/ms-excel.svg)
> ![](resources/svg/ms-word.svg)
> ![](resources/svg/ms-outlook.svg)
> ![](resources/svg/ms-access.svg)
> **and any** ![](resources/svg/ms-office.svg) **VBA host**

> [!IMPORTANT]
> **Platform**
>
> ![](resources/svg/windows.svg)
> Currently available only on **Windows**, as it relies on Windows-specific APIs (`ws2_32`, `secur32`, `crypt32`, `advapi32`, `kernel32`).

## Table of Contents

- [What is Wasabi](#what-is-wasabi)
- [Quick Start](#quick-start)
- [Architecture](#architecture)
  - [Why a Standard Module](#why-a-standard-module-bas-instead-of-classes-cls)
  - [Assembly Engine](#the-assembly-engine-thunks)
  - [Modular Design](#modular-design-dumb-pipe--extensions)
  - [Execution Model](#execution-model-single-thread-polling-and-async)
- [Features](#features)
  - [Middleware Pipeline](#the-middleware-pipeline)
  - [Protocol and Compression Extensions](#protocol-and-compression-extensions)
- [Examples](#examples)
- [Performance](#performance)
- [Compatibility](#compatibility)
- [Use Cases](#use-cases)
- [Roadmap](#roadmap)
- [Community & Acknowledgements](#community--acknowledgements)
- [Contributing](#contributing)
- [Security](#security)
- [License](#license)

## What is Wasabi

Wasabi is a single, self-contained `.bas` module that brings real-time networking to any Microsoft Office VBA host. It was designed to feel familiar to developers who have worked with [socket.io](https://socket.io) in Node.js, but runs entirely inside the Office ecosystem with no external runtimes, no COM registration, and no installers.

A single file drop into any VBA project is all it takes. No references need to be enabled in **Tools -> References**.

Beyond WebSocket, Wasabi ships a full MQTT client with MQTT 5 extensions (User Properties, Reason Codes, metadata handling), a first-class raw TCP client, NTLM/Kerberos proxy authentication, RTT latency measurement, fine-grained TLS certificate control, a composable middleware pipeline, a pluggable compression architecture, and a `WSAAsyncSelect`-based event-driven async system that fires callbacks while Excel is idle without any polling loop. The module compiles cleanly on 32-bit and 64-bit Office hosts, from Windows XP to Windows 11, through conditional compilation (`#If VBA7`).

## Quick Start

### Import

[Download the latest release](../../releases) and import `Wasabi.bas` into your VBA project via **File -> Import File** in the VBA editor.

### Connect and Send a Message

```vb
Dim h As Long

If WebSocketConnect("wss://echo.websocket.org", h) Then
    WebSocketSend "Hello, Wasabi!", h

    Dim msg As String
    msg = WebSocketReceive(h)

    If msg <> "" Then
        Debug.Print "Received: " & msg
    End If

    WebSocketDisconnect h
End If
```

### Connect with TLS Certificate Validation

```vb
Dim h As Long

WebSocketSetCertValidation True, h
WebSocketSetRevocationCheck True, h

If WebSocketConnect("wss://example.com/ws", h) Then
    WebSocketSend "Secure hello", h
    WebSocketDisconnect h
End If
```

### MQTT with QoS 2 and MQTT 5 User Properties

```vb
Dim h As Long

WebSocketConnect "wss://broker.hivemq.com:8443/mqtt", h, , , "mqtt"
WebSocketSetOfflineQueueing True, h

MqttConnect "WasabiClient_123", , , 60, h

' Publish with QoS 2 and attach a user property (MQTT 5 metadata)
MqttPublish "sensors/data", "Value: 42", 2, False, "source", "wasabi", h
```

### Connect Through a Proxy

```vb
Dim h As Long

' Hardcoded proxy or use WebSocketAutoDiscoverProxy()
WebSocketSetProxy "proxy.company.com", 8080, "user", "pass", 0, h
WebSocketSetProxyNtlm True, h

If WebSocketConnect("wss://example.com/ws", h) Then
    WebSocketSend "Behind the firewall", h
    WebSocketDisconnect h
End If
```

### Auto-Reconnect with Keepalive and Ping Jitter

```vb
Dim h As Long

WebSocketSetAutoReconnect True, 5, 1000, h
WebSocketSetPingInterval 30000, 5000, h  ' 30s interval, up to 5s of random jitter

If WebSocketConnect("wss://example.com/ws", h) Then
    Do While WebSocketIsConnected(h)
        Dim msg As String
        msg = WebSocketReceive(h)
        If msg <> "" Then Debug.Print "Received: " & msg
        DoEvents
    Loop
End If
```

### Event-Driven Async

Instead of polling in a loop, register a handler object and let Wasabi call your callbacks automatically while Excel is idle.

```vb
' In a standard module at module level
Private g_Handler As cMyHandler
Private g_Handle As Long

Public Sub StartAsync()
    Set g_Handler = New cMyHandler
    If WebSocketConnect("wss://stream.binance.com:9443/ws/btcusdt@trade", g_Handle) Then
        WasabiUseAsync g_Handler, g_Handle
        ' Sub ends here. Callbacks fire while Excel is idle.
    End If
End Sub
```

Handler class (`cMyHandler`):

```vb
Option Explicit

Public Sub OnConnect(ByVal handle As Long)
    Debug.Print "Connected. Handle: " & handle
End Sub

Public Sub OnReceive(ByVal handle As Long)
    Dim msg As String
    Do While WebSocketGetPendingCount(handle) > 0
        msg = WebSocketReceive(handle)
        Debug.Print "Received: " & msg
    Loop
End Sub

Public Sub OnReadyToSend(ByVal handle As Long)
End Sub

Public Sub OnClose(ByVal handle As Long)
    Debug.Print "Connection closed. Handle: " & handle
End Sub

Public Sub OnError(ByVal handle As Long, ByVal errorCode As Long, ByVal eventType As Long)
    Debug.Print "Error on handle " & handle & " | Code: " & errorCode
End Sub
```

> [!IMPORTANT]
> The handler object must be stored at module or workbook level, never as a local variable inside a `Sub`. Always call `WebSocketDisconnect` before resetting the VBA project or closing the workbook.

### Pluggable Compression

```vb
Dim h As Long

' Register any class implementing Deflate/Inflate
Dim deflate As New ExtWasabiZlib
WasabiUseCompression deflate, h

If WebSocketConnect("wss://example.com/ws", h) Then
    Debug.Print "Compression active: " & WebSocketGetDeflateEnabled(h)
    WebSocketSend "Compressed payload", h
    WebSocketDisconnect h
End If
```

> [!NOTE]
> Compression is fully opt-in and algorithm-agnostic. If no handler is registered, the connection proceeds normally without compression. The `permessage-deflate` reference implementation (`ExtWasabiZlib.cls`) is a separate extension. Documentation and setup instructions are provided alongside that extension by whoever ships it.

### Using the Middleware Pipeline

```vb
' MyLogger.cls
Public Sub OnBeforeSend(ByVal handle As Long, ByRef data() As Byte)
    Debug.Print "[OUT] " & UBound(data) + 1 & " bytes"
End Sub

Public Sub OnAfterReceive(ByVal handle As Long, ByRef data() As Byte)
    Debug.Print "[IN]  " & UBound(data) + 1 & " bytes"
End Sub

Public Sub OnConnect(ByVal handle As Long): End Sub
Public Sub OnDisconnect(ByVal handle As Long): End Sub
```

```vb
Dim h As Long
Dim logger As New MyLogger

WasabiUseMiddleware logger, h

If WebSocketConnect("wss://example.com/ws", h) Then
    WebSocketSend "Instrumented message", h
    WebSocketDisconnect h
End If
```

### Raw TCP (plain and TLS)

```vb
' Plain TCP with delimiter-based read
Dim h As Long

If TcpConnect("tcpbin.com", 4242, h) Then
    TcpSendText "hello" & vbCrLf, h

    Dim line As String
    line = TcpReceiveUntil(vbCrLf, 3000, h)
    Debug.Print "Echo: " & line

    TcpDisconnect h
End If
```

```vb
' TCP with TLS (full Schannel stack)
Dim h As Long

If TcpConnectTLS("example.com", 443, h) Then
    TcpSendText "GET / HTTP/1.0" & vbCrLf & "Host: example.com" & vbCrLf & vbCrLf, h

    Dim t As Long, msg As String
    t = GetTickCount()
    Do While TickDiff(t, GetTickCount()) < 5000
        msg = TcpReceiveText(h)
        If Len(msg) > 0 Then Exit Do
        DoEvents
    Loop

    Debug.Print Left(msg, 200)
    TcpDisconnect h
End If
```

For the complete API reference with all parameters, return values, and usage notes, see the [API Reference](docs/API_REFERENCE.md).

## Architecture

### Why a Standard Module (.bas) instead of Classes (.cls)?

This is a deliberate design choice, not a limitation. Several concrete constraints of the VBA environment make a procedural standard module the right tool for networking at this level.

**Zero COM overhead.** Every VBA class is a COM object. Instantiating it, dispatching method calls through `IDispatch`, and managing reference counts all add overhead that accumulates when processing frames at high frequency. A standard `.bas` module communicates directly with the CPU and memory, bypassing the COM layer entirely.

**Static connection pool and data-oriented design.** Wasabi manages up to 64 concurrent connections using a statically allocated pool of User-Defined Types (`WasabiConnection`). All connection state lives in a contiguous block of memory rather than objects scattered across the heap. This is more cache-friendly and eliminates heap fragmentation in long-running Office sessions.

**Native Win32 API alignment.** Low-level networking requires heavy use of memory pointers (`StrPtr`, `VarPtr`) and direct memory manipulation (`RtlMoveMemory`). Standard modules provide a flatter memory model for passing data to Windows Kernel and Security APIs. Passing class properties to Win32 APIs often requires temporary copies; procedural modules allow in-place processing, which is essential for the zero-copy receive model.

**Minimal integration friction.** State is managed globally through a simple integer handle. There are no object lifecycles to manage and no risk of a variable falling out of scope and silently terminating a live connection.

**Developer experience.** The global scope of standard modules exposes public enums (`WasabiState`, `WasabiConnectionMode`, `WasabiError`) and types (`WasabiStats`) everywhere in the project without instantiating anything, with full IntelliSense support. Functions like `WebSocketGetLastError` return specific `WasabiError` enum values, enabling clean `Select Case` blocks and self-documenting error handling.

> [!TIP]
> This architecture transforms Wasabi from a simple script into a high-performance networking engine, bringing C-level memory management and stability to the VBA ecosystem.

### The Assembly Engine (Thunks)

<div align="center">
  <img src="resources/using-assembly.png" alt="Wasabi Assembly Engine" />
</div>

VBA is an interpreted language that excels at COM automation but is inherently slow at processing large byte arrays sequentially, because every loop iteration carries bounds-checking and variant type conversion overhead. For operations like masking WebSocket frames or scanning megabyte-sized TCP buffers for delimiters, this is unacceptable.

Wasabi solves this with **safe thunks**: compiled machine code (x86 and x64) injected into executable memory at runtime and invoked through `CallWindowProcW`.

**How it works:**

1. On initialization, `VirtualAlloc` requests a block of memory with `PAGE_EXECUTE_READWRITE` permissions.
2. The raw opcode bytes for each thunk are copied into that block via `RtlMoveMemory`.
3. When a heavy byte-level operation is needed, `CallWindowProcW` executes the block as if it were a native C function.

**The four thunks Wasabi ships:**

`ws_mask` applies the mandatory WebSocket XOR mask (RFC 6455) to outbound frames. It eliminates the VBA loop bottleneck entirely, allowing megabytes of payload to be masked in microseconds.

`mem_zero` is a hardware-level memory wipe using the `rep stosb` instruction. Called during connection cleanup to securely erase decrypted payloads, TLS buffers, and proxy credentials from RAM.

`mem_find` is hardware-accelerated byte-pattern matching using `repe cmpsb`. It locates delimiters such as `\r\n` inside massive TCP buffers instantly. This thunk is also the engine behind `TcpReceiveUntil`.

`tick_diff` is a lightweight counter helper used for timeout and latency calculations, handling `GetTickCount` overflow correctly.

**Calling conventions:** x64 uses the Microsoft x64 convention (arguments in `RCX`, `RDX`, `R8`, `R9`). x86 uses `stdcall` (arguments pushed to the stack, cleaned with `ret 16`).

The original NASM/FASM source for all thunks, with full comments and compilation instructions, is in the [`dev/asm`](dev/asm) directory.

### Modular Design: Dumb Pipe + Extensions

Wasabi is built on strict separation of concerns across four layers.

**Raw Transport (Dumb Pipe).** Manages TCP sockets, TLS via Schannel, Happy Eyeballs (RFC 6555), and proxy authentication. Treats the connection as a raw byte stream with no knowledge of any protocol.

**Protocol Layer.** Injected via `WasabiUseProtocol`. Handles WebSocket framing, MQTT message parsing, or any custom binary or text protocol. The engine dispatches parsed messages to the registered handler. MQTT support is built-in and uses this slot internally.

**Middleware Layer.** Registered via `WasabiUseMiddleware`. Intercepts every byte flowing inbound and outbound. Multiple middleware objects are chained in registration order. Ideal for logging, encryption, HMAC signing, or metrics collection without touching application code.

**Compression Layer.** Registered via `WasabiUseCompression`. Any class implementing `Deflate` and `Inflate` methods can be plugged in. The core module has no dependency on any compression library and only calls this interface if a handler is registered.

This separation means security patches, protocol additions, and algorithm changes are confined to isolated components, without ever requiring a fork of the main module.

### Execution Model: Single-Thread, Polling, and Async

VBA is single-threaded. One execution thread is shared between your code and the Office UI, so there is no native background socket listener.

**Polling model.** Incoming messages accumulate in an internal ring buffer (up to 512 messages) and are delivered when you call `WebSocketReceive`. Each call runs keepalive maintenance (pings, inactivity timeout, MTU probes), checks the OS socket buffer via `FIONREAD`, reads available data, decrypts if TLS is active, parses WebSocket frames, runs all registered inbound middleware, and returns the oldest queued message. Between calls, the kernel continues buffering incoming data. No messages are lost and the connection does not drop regardless of polling frequency. Simple send/receive/disconnect workflows work fine without any loop. For live dashboards and reactive scenarios, the recommended polling pattern is `Application.OnTime`.

**Async model.** Wasabi also supports a fully event-driven mode via `WasabiUseAsync`. Under the hood it uses `WSAAsyncSelect` to register the socket with a hidden Win32 window, and a native machine-code thunk subclasses that window to dispatch `FD_READ`, `FD_WRITE`, `FD_CLOSE`, and `FD_CONNECT` events to your handler object. Callbacks fire while Excel is idle without any polling loop on your part. If your code is running a tight loop without `DoEvents`, messages queue in the Windows message pump and are delivered as soon as your code finishes or yields. The async thunk includes a VBA runtime liveness guard that falls back to `DefWindowProcW` if the project has been reset, preventing crashes. See the [API Reference](docs/API_REFERENCE.md) for the full async callback contract and usage notes.

## Features

### The Middleware Pipeline

Any VBA class implementing the following four methods can be registered as middleware:

```vb
Public Sub OnBeforeSend(ByVal handle As Long, ByRef data() As Byte)
Public Sub OnAfterReceive(ByVal handle As Long, ByRef data() As Byte)
Public Sub OnConnect(ByVal handle As Long)
Public Sub OnDisconnect(ByVal handle As Long)
```

Multiple middleware objects can be chained on the same handle and are executed in registration order for both inbound and outbound data:

```vb
Dim logger As New MyLoggingMiddleware
Dim encryptor As New MyEncryptionMiddleware

WasabiUseMiddleware logger, h
WasabiUseMiddleware encryptor, h
```

Typical uses include transparent byte logging, custom payload encryption or HMAC signing before frames leave the socket, custom compression schemes, and byte counters or message-rate metrics without polluting application code.

### Protocol and Compression Extensions

**Protocol Handler (`WasabiUseProtocol`).** A single object that receives parsed text and binary messages after the WebSocket frame layer has processed them. This is the cleanest integration point for application-level protocols without touching the transport engine.

```vb
Dim myProto As New MyMqttProtocol
WasabiUseProtocol myProto, h
```

**Compression Handler (`WasabiUseCompression`).** An object implementing `Deflate` and `Inflate`. This slot replaces the built-in compression path, allowing any algorithm such as LZ4, Brotli, or Zstandard to be supplied without modifying the core module.

```vb
Dim lz4 As New MyLZ4Compressor
WasabiUseCompression lz4, h
```

Both extensions receive `OnConnect` and `OnDisconnect` lifecycle callbacks automatically.

## Examples

Ready-to-run `.xlsm` workbooks are in the [`examples/`](examples/) folder. Highlights include:

**Crypto Live Ticker.** Connect to public exchange streams such as Binance and update cells in real time.

**MQTT QoS 2 Dashboard.** Full-duplex IoT dashboard with guaranteed message delivery, MQTT 5 User Properties, ping jitter, and offline queueing.

**Non-Blocking UI (Event Loop).** The definitive architectural pattern using `Application.OnTime` to keep spreadsheets 100% interactive while listening to background data.

**Async Event-Driven.** Register a handler class and receive callbacks while Excel is idle, with no polling loop required.

**High-Speed Batching and Corporate Auth.** Advanced configurations for strict TLS, system proxies, and high-throughput telemetry.

**Middleware Pipeline.** How to inject logging, encryption, or transformation layers into the data flow without touching core send/receive logic.

[Explore the Examples Suite](examples/)

## Performance

All cryptographic and encoding primitives are delegated to native Windows APIs (`advapi32.dll`, `crypt32.dll`), and intensive byte processing is routed directly to the CPU via the Assembly Engine. This yields throughput close to the hardware limit even inside the interpreted VBA runtime.

![Throughput Benchmark](resources/benchmark-throughput.png)

> [!NOTE]
> SHA-1 now runs at **182,8 MB/s** (down from 1.42 s per 128 KB in pure VBA). Base64 operations stay around **41 MB/s**, UTF-8 conversion exceeds **1 GB/s**, and WebSocket frame construction tops **25 MB/s**. The test harness and raw data are in [`benchmark/`](benchmark/).

### Stress Test Results

The engine was subjected to a continuous, single-threaded stress test handling over 10 MB of concurrent traffic, cryptographic masking, and hardware-level buffer scanning, all on a standard Office VBA thread.

| Transport Layer | Payload / Operation | Execution Time | Throughput / Notes |
| :--- | :--- | :--- | :--- |
| **TCP TLS (Schannel)** | DNS + mTLS Handshake + HTTP GET | `656 ms` | Native C-level latency (Happy Eyeballs IPv6) |
| **TCP Raw** | 10x `TcpReceiveUntil` (`\r\n`) buffer scans | `5156 ms` | Instant `mem_find` hardware delimiter scan |
| **WSS Deflate** | 100x 100KB payloads (**10 MB Total**) | `4469 ms` | **~2.2 MB/s** with inline ASM XOR masking |
| **MQTT 5 (QoS 2)** | Subscribe + 50x QoS 2 Publishes | `1219 ms` | Flawless In-Flight queue & metadata handling |

## Compatibility

Wasabi uses only native Windows DLLs present on every version of Windows since XP: `ws2_32.dll`, `secur32.dll`, `kernel32.dll`, `advapi32.dll`, and `crypt32.dll`. No third-party installers, no COM registration, no ActiveX controls, no Python runtime, no .NET packages. Dropping the `.bas` file into a VBA project is all it takes.

Many competing modules depend on `WinHttpWebSocket*` functions introduced in Windows 8, which causes silent failures on Windows 7 machines that remain common in corporate and industrial environments. Wasabi has no such limitation.

### Operating System

| Version | Support |
|---|---|
| ![](resources/svg/windows.svg) Windows XP | ![](resources/svg/checked.svg) |
| ![](resources/svg/windows.svg) Windows Vista | ![](resources/svg/checked.svg) |
| ![](resources/svg/windows.svg) Windows 7 | ![](resources/svg/checked.svg) |
| ![](resources/svg/windows.svg) Windows 8 / 8.1 | ![](resources/svg/checked.svg) |
| ![](resources/svg/windows.svg) Windows 10 | ![](resources/svg/checked.svg) |
| ![](resources/svg/windows.svg) Windows 11 | ![](resources/svg/checked.svg) |

### Office and VBA Host

| Environment | Support |
|---|---|
| ![](resources/svg/ms-excel.svg) Excel 32-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-excel.svg) Excel 64-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-word.svg) Word 32-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-word.svg) Word 64-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-powerpoint.svg) PowerPoint 32-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-powerpoint.svg) PowerPoint 64-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-access.svg) Access 32-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-access.svg) Access 64-bit | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-office.svg) Any VBA7 host (Office 2010+) | ![](resources/svg/checked.svg) |
| ![](resources/svg/ms-office.svg) VBA6 (Office 2007 and earlier) | ![](resources/svg/checked.svg) |

32-bit and 64-bit compatibility is handled transparently through `#If VBA7` conditional compilation across all API declarations. The same `.bas` file compiles correctly on Office 2007 32-bit and Office 365 64-bit on Windows 11.

## Use Cases

**Bots and chat integrations.** Connect to Discord, Slack, or Telegram gateways and handle real-time events directly from Excel or Word.

**Trading and finance.** Stream live market data from exchanges like Binance, Coinbase, or B3 into spreadsheet cells with millisecond-level latency.

**Live dashboards.** Update cells in real time without manual refreshes or HTTP polling.

**IoT and industrial SCADA.** Receive sensor data from ESP32, Raspberry Pi, or PLC systems via WebSocket or MQTT natively into Office.

**Corporate automation.** Connect Office to internal WebSocket APIs behind firewalls and NTLM/Kerberos proxies without requiring IT to install third-party software.

**Raw TCP automation.** Communicate directly with legacy PLC systems, industrial equipment, SMTP servers, or custom binary protocols that do not speak WebSocket or HTTP.

**Custom protocol engines.** Use the Protocol Handler and Middleware slots to implement proprietary binary protocols cleanly on top of the Wasabi transport layer.

## Roadmap

### ![](resources/svg/completed.svg) Completed

- [x] IPv6 and SNI support
- [x] Mutual TLS (mTLS) for client certificate authentication via PFX or system certificate store
- [x] SOCKS5 proxy support
- [x] HTTP/2 upgrade via ALPN (opt-in)
- [x] NTLM/Kerberos authentication for HTTP proxies
- [x] **Windows System Proxy Auto-Discovery** (PAC scripts via `winhttp.dll`)
- [x] MQTT client (3.1.1 with **MQTT 5 extensions**): QoS 1 and **QoS 2 (Exactly Once)**, User Properties (`metaKey`/`metaValue`), Reason Codes, In-Flight queue with Packet ID tracking (PUBREC/PUBREL/PUBCOMP)
- [x] RTT latency measurement (`GetLatency`)
- [x] `permessage-deflate` compression (RFC 7692) via pluggable extension
- [x] Zero-copy receive buffers and MTU-aware frame sizing
- [x] Send batching (text and binary)
- [x] Close frame payload parsing (code and reason)
- [x] Happy Eyeballs (RFC 6555)
- [x] Configurable CRL/OCSP certificate revocation checking
- [x] **Strict State Machine** (`STATE_CONNECTING`, `STATE_OPEN`, `STATE_CLOSING`, `STATE_CLOSED`)
- [x] **Offline Queueing**: messages sent during a disconnect are buffered and flushed on reconnect
- [x] **Ping Jitter**: pseudo-random variance on keepalive pings to avoid strict gateway filters
- [x] **Native TCP Client** (`TcpConnect`, `TcpConnectTLS`) sharing the full Schannel engine
- [x] TCP MTU discovery, `NoDelay`, inactivity timeout, and proxy support
- [x] **`TcpBroadcast` / `TcpBroadcastText`**: send to all active TCP connections simultaneously
- [x] **`TcpReceiveUntil`**: delimiter-based blocking read
- [x] **Middleware Pipeline** (`WasabiUseMiddleware`)
- [x] **Protocol Handler** (`WasabiUseProtocol`) for pluggable application protocol injection
- [x] **Compression Handler** (`WasabiUseCompression`) for decoupled compression extensions
- [x] **Core Refactoring**: raw transport decoupled from protocol logic
- [x] **Modular Compression**: `zlib1.dll` dependency isolated into the `ExtWasabiZlib.cls` extension
- [x] **Async Event-Driven Mode** (`WasabiUseAsync`): `WSAAsyncSelect`-based socket notifications with a native thunk dispatcher, eliminating polling loops

### ![](resources/svg/looking.svg) In Progress

Nothing currently in progress. Have an idea? [Open an issue](../../issues) or a [pull request](../../pulls).

## Community & Acknowledgements

Wasabi builds on years of experimentation and innovation from the VBA/VB6 low-level networking community.

Special thanks to the developers and projects that helped inspire parts of Wasabi's architecture and engineering direction:

- <img src="resources/ms-vba.png" width="18" /> [vba-websocket](https://github.com/EagleAglow/vba-websocket): WebSocket client implementation for VBA by <img src="https://github.com/EagleAglow.png" width="18" /> [**EagleAglow**](https://github.com/EagleAglow)
- <img src="resources/ms-vba.png" width="18" /> [VbAsyncSocket](https://github.com/wqweto/VbAsyncSocket): High-performance asynchronous socket library for VB6/VBA by <img src="https://github.com/wqweto.png" width="18" /> [**wqweto**](https://github.com/wqweto)
- <img src="resources/ms-vba.png" width="18" /> [stdVBA](https://github.com/sancarn/stdVBA): Modern standard library and utilities for VBA by <img src="https://github.com/sancarn.png" width="18" /> [**sancarn**](https://github.com/sancarn)
- <img src="resources/ms-vba.png" width="18" /> [TlsSocketWSS-vb6](https://github.com/Maatooh/TlsSocketWSS-vb6): TLS-enabled WebSocket implementation for VB6 by <img src="https://github.com/Maatooh.png" width="18" /> [**Maatooh**](https://github.com/Maatooh)
- <img src="resources/ms-vba.png" width="18" /> [vb6-websocket-server-ssl](https://github.com/JoshyFrancis/vb6-websocket-server-ssl): Secure WebSocket server for VB6 by <img src="https://github.com/JoshyFrancis.png" width="18" /> [**JoshyFrancis**](https://github.com/JoshyFrancis)
- <img src="resources/ms-vba.png" width="18" /> [VBA_WinsockAPI_TCP_Sample](https://github.com/papanda925/VBA_WinsockAPI_TCP_Sample): WinSock TCP examples and networking experiments for VBA by <img src="https://github.com/papanda925.png" width="18" /> [**papanda925**](https://github.com/papanda925)
- <img src="resources/ms-vba.png" width="18" /> [VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools): Native Memory Manipulation in VBA by <img src="https://github.com/cristianbuse.png" width="18" /> [**cristianbuse**](https://github.com/cristianbuse)

> [!NOTE]
> We are honored that several developers from the VBA/VB6 networking community have recognized and starred this repository. Their work helped shape the ecosystem that made Wasabi possible.

## Stargazers over time

[![Stargazers over time](https://starchart.cc/uesleibros/wasabi.svg?variant=adaptive)](https://github.com/uesleibros/wasabi/stargazers)

## Contributing

Bug reports, feature requests, and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## Security

Do **not** report vulnerabilities through public issues. See [SECURITY.md](SECURITY.md) and use GitHub Private Vulnerability Reporting.

## License

**MIT**, free for personal and commercial use. See [LICENSE](LICENSE).
