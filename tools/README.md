# Tools

This folder contains auxiliary tools for developing and testing Wasabi.
None of these tools are required to use the module they exist purely
for contributors and advanced users who want to run local tests.

## Contents

| File / subfolder | Description |
|:---|:---|
| `wasabi-test-server.c` | Lightweight WebSocket echo server written in C. Supports plain TCP (`ws://`) and TLS (`wss://`) via OpenSSL. Used for local integration tests. |

## wasabi-test-server

A single‑file WebSocket echo server that handles the WebSocket handshake,
echoes text and binary frames, responds to pings with pongs, and performs
a proper close handshake.

> [!NOTE]
> This server is intended exclusively for local testing. It accepts only
> one connection at a time and has no authentication or rate limiting.

### Compilation

**Linux / Termux**

```bash
gcc -O2 -o wasabi-test-server wasabi-test-server.c -lssl -lcrypto
```

**Ubuntu / Debian** — install OpenSSL development headers first:

```bash
sudo apt install libssl-dev
gcc -O2 -o wasabi-test-server wasabi-test-server.c -lssl -lcrypto
```

**Windows (MinGW‑w64 cross‑compilation)**

```bash
x86_64-w64-mingw32-gcc -O2 -o wasabi-test-server.exe wasabi-test-server.c -lssl -lcrypto -lws2_32
```

### Usage

**Plain WebSocket (port 9001)**

```bash
./wasabi-test-server 9001
```

**TLS WebSocket (port 9443 with test certificates)**

```bash
./wasabi-test-server 9443 ../resources/certs/server.pem ../resources/certs/server.key
```

> [!TIP]
> Use the TLS mode together with the certificates in `resources/certs/`
> to validate Wasabi’s `wss://` and mTLS functionality in a local
> environment.

### Testing with Wasabi

```vb
' Plain WebSocket echo
Dim h As Long
If WebSocketConnect("ws://localhost:9001", h) Then
    WebSocketSend "Hello from Wasabi", h
    Debug.Print WebSocketReceive(h)   ' prints "Hello from Wasabi"
    WebSocketDisconnect h
End If
```
