# permessage-deflate Negotiation

## Client offer

When `DeflateEnabled = True` (user parameter in `WebSocketConnect` or
`WebSocketSetDeflate`), Wasabi adds a `Sec-WebSocket-Extensions` header to
the handshake request:

    Sec-WebSocket-Extensions: permessage-deflate; client_no_context_takeover; server_no_context_takeover; client_max_window_bits=12

The exact parameters depend on configured options:

- `DeflateContextTakeover = False` → `client_no_context_takeover`
- `InflateContextTakeover = False` → `server_no_context_takeover`
- `ClientMaxWindowBits <> 15` → `client_max_window_bits=<value>`

## Server response parsing

`ParseDeflateResponse(handle, response)` scans the server's
`Sec-WebSocket-Extensions` header and extracts the negotiated parameters:

- If the header is missing, `DeflateEnabled` becomes `False` (server does
  not support the extension).
- `client_no_context_takeover` → `DeflateContextTakeover = False`
- `server_no_context_takeover` → `InflateContextTakeover = False`
- `server_max_window_bits=<value>` → sets `DeflateWindowBits = -value`,
  stores `ServerMaxWindowBits`.
- `client_max_window_bits=<value>` → sets `InflateWindowBits = -value`,
  stores `ClientMaxWindowBits`.

After successful parsing, `DeflateActive = True`, enabling compression
for subsequent frames.

## Reference

- [RFC 7692](https://datatracker.ietf.org/doc/html/rfc7692)
