# Optional zlib Dependency (permessage-deflate)

Wasabi supports the WebSocket `permessage-deflate` extension (RFC 7692),
which compresses message payloads to reduce bandwidth usage.

This feature relies on **zlib**, a widely‑used, open‑source compression
library. To keep the core module dependency‑free, zlib is **not bundled**
with Wasabi and must be provided by the user only if compression is desired.

## Why zlib is optional

- Without zlib, Wasabi remains 100% self‑contained. Every other feature
  (TLS, proxies, MQTT, etc.) works exactly as before.
- When zlib is available, Wasabi automatically negotiates
  `permessage-deflate` during the WebSocket handshake.
- If zlib is missing, compression is silently disabled. The connection
  proceeds without compression, and a diagnostic message is logged.

## Obtaining zlib

You need the correct build of zlib for your Office environment.

| Office architecture | Required DLL |
|---|---|
| 32‑bit | `zlib1_x86.dll` |
| 64‑bit | `zlib1_x64.dll` |

Pre‑compiled DLLs can be downloaded from the official zlib website:

- [https://zlib.net](https://zlib.net)

> [!NOTE]
> Many applications already ship a `zlib1.dll` in their installation
> directory. Wasabi will attempt to load a generic `zlib1.dll` as a
> fallback, but using the architecture‑specific name is preferred.

## Where to place the DLL

Wasabi searches for zlib in the following locations, in order:

1. The folder containing the host document (Excel workbook, Word document, etc.)
2. Subdirectories of the host document: `\lib`, `\deps`, `\dlls`, `\zlib`, `\bin`, `\x64`, `\x86`, `\native`
3. The Windows system directories (`System32` and `SysWOW64`)
4. Any directory listed in the `PATH` environment variable

Within each location, Wasabi looks first for the architecture‑specific name
(`zlib1_x64.dll` or `zlib1_x86.dll`), then for the generic `zlib1.dll`.

> [!TIP]
> The simplest setup is to drop the DLL into the same folder as your Excel
> file. No registration or installation is required.

## How Wasabi loads zlib

The loading process is handled by three private functions inside `Wasabi.bas`:

### `GetZlibName()`

Returns the architecture‑appropriate filename.

### `FindZlibPath()`

Iterates through a list of known directories and returns the first one that
contains a matching DLL.

### `LoadZlib()`

Uses the Windows `LoadLibrary` API to load the DLL. If none of the search
paths yield a valid library, the function logs a warning and exits. The
result is cached so that the search runs only once per session.

## Enabling compression

Compression is requested on a per‑connection basis:

```vb
Dim h As Long
WebSocketConnect "wss://example.com/ws", h, True, True
```

The third parameter enables `permessage-deflate`. If zlib was loaded
successfully, the extension is included in the handshake. The fourth
parameter controls context takeover.

To enable compression on an existing (disconnected) handle:

```vb
WebSocketSetDeflate True, True, h
```

To check whether compression is active after the handshake:

```vb
If WebSocketGetDeflateEnabled(h) Then
    Debug.Print "Compression is active"
End If
```

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `DeflateEnabled` stays `False` after connecting | The server does not support `permessage-deflate` (this is normal for many public servers) |
| `DeflateEnabled` stays `False` even with a compatible server | zlib DLL was not found; check the log for `LoadZlib: zlib1.dll not found` |
| Connection fails with `ERR_COMPRESSION` | (future) The server and client could not agree on deflate parameters |

> [!NOTE]
> `permessage-deflate` reduces bandwidth usage, not latency. For small
> messages (e.g., typical WebSocket JSON frames), the compression overhead
> may slightly increase CPU usage without noticeable bandwidth savings.
> It is most beneficial for large payloads or connections with limited
> bandwidth.
