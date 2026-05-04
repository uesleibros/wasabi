# Happy Eyeballs Algorithm in Wasabi

## Steps in `ResolveAndConnect`

1. **DNS resolution**: `sock_getaddrinfo` retrieves both AF_INET6 and
   AF_INET addresses. If `getaddrinfo` fails, falls back to
   `gethostbyname` (IPv4 only).

2. **IPv6 starts first**: An IPv6 socket is created and set to
   non‑blocking (`FIONBIO`). `connect()` is called immediately.

3. **250ms delay for IPv4**: If both address families are available, the
   code waits up to `HAPPY_EYEBALLS_DELAY_MS` (250ms) for the IPv6
   socket to connect. During this wait it repeatedly calls `select()`
   with a 50ms timeout, checking whether the socket became writable
   (connection succeeded) or had an exception (failed).

4. **IPv4 race**: After the delay (or if IPv6 fails early), an IPv4
   socket is created and connects in parallel.

5. **Final selection**: Both sockets race with a total timeout of 10s.
   The first to connect is kept; the other is closed.

## Why non‑blocking sockets?

VBA cannot use `WSAAsyncSelect` or IOCP. Non‑blocking mode with
`select()` is the only portable way to check connection completion
without blocking the UI.

## Why 250ms?

RFC 8305 recommends 250ms as the initial delay before starting IPv4.
This gives IPv6 a head start but falls back quickly if IPv6 is broken.
