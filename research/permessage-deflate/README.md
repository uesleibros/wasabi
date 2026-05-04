# permessage-deflate Implementation

This directory documents Wasabi's implementation of the WebSocket
`permessage-deflate` extension (RFC 7692), which compresses message
payloads using raw deflate (zlib).

## Related

- `../zlib-stdcall/` – How the required stdcall zlib DLL was obtained.
- `../docs/README.md` – User-facing documentation for compression settings.
