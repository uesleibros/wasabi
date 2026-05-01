# Package

This folder contains **Wasabi.bas**, the single, self‑contained file that
brings the entire WebSocket stack to any VBA project.

## What it is

* A **plain text** `.bas` module exported from the VBA IDE.
* **Zero dependencies** — no DLLs, no COM registrations, no installers.
* **Self‑contained** — all the logic lives inside this one file.

When you import `Wasabi.bas`, the VBA runtime compiles it on the spot.
There is no build step, no binary, and no packaging. What you see is
exactly what runs inside the Office process.

## What’s inside

Internally, `Wasabi.bas` is organised in several distinct layers that
work together to deliver a complete, production‑grade WebSocket client:

| Layer | What it does |
|:---|:---|
| **Public API** | High‑level functions (`WebSocketConnect`, `Send`, `Receive`, etc.) designed for everyday use. |
| **WebSocket Protocol** | Frame construction, masking, fragmentation, ping / pong, close handshake. |
| **TLS Security** | Native Schannel SSPI implementation for `wss://` (TLS 1.2 & 1.3). |
| **Transport** | Raw Winsock2 TCP sockets with Happy Eyeballs (IPv6 + IPv4) and proxy support (HTTP / SOCKS5). |
| **Windows Declarations** | All used API functions (`ws2_32`, `secur32`, `kernel32`, `advapi32`, `crypt32`) are declared at the top of the module and guarded by `#If VBA7` for automatic 32‑bit / 64‑bit compatibility. |

## Compatibility

* **Windows** — XP, Vista, 7, 8, 10, 11 (x86 and x64)
* **Office** — 2007 to 365 (32‑bit and 64‑bit)
* **Hosts** — Excel, Word, PowerPoint, Access, and any VBA‑enabled application

Every API declaration uses `#If VBA7` to switch between `LongPtr`/`PtrSafe`
and classic `Long`/`Declare`, so the same file works everywhere without
modification.

## How to use

1. In the VBA editor, click **File → Import File…**
2. Select `Wasabi.bas` from this folder.
3. No additional steps are required — no references, no tools, no setup.

After importing, you can call `WebSocketConnect` directly from any module.

```vb
Dim h As Long
If WebSocketConnect("wss://echo.websocket.org", h) Then
    WebSocketSend "Hello from Wasabi", h
    Debug.Print WebSocketReceive(h)
    WebSocketDisconnect h
End If
```

> [!TIP]
> The complete API reference is available in [`docs/API_REFERENCE.md`](../docs/API_REFERENCE.md).
