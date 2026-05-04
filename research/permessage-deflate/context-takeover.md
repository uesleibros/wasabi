# Context Takeover

`permessage-deflate` allows the LZ77 sliding window to be preserved
between messages (context takeover), improving compression ratios.
However, it consumes more memory (both client and server must keep the
sliding window) and can affect multiplexing.

## Wasabi defaults

- `DeflateContextTakeover = True` (client retains deflate window)
- `InflateContextTakeover = True` (server retains inflate window)

## When disabled

If the client offers `client_no_context_takeover`, the deflate stream is
ended after each message (`DeflateReady = False`), and a fresh stream is
created for the next message.

## Memory impact

A 32 KB sliding window is maintained per stream. With context takeover,
this memory is retained for the lifetime of the connection.

## Reference

- RFC 7692, Section 5.1
