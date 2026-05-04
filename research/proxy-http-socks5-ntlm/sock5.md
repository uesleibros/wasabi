# SOCKS5 Proxy

`DoProxySOCKS5(handle)` implements the SOCKS5 handshake (RFC 1928).

1. **Greeting**: Sends supported authentication methods (0x00 = no auth,
   0x02 = user/password). Receives server choice.

2. **Authentication** (if required): Sends username and password
   (method 0x02). Receives status.

3. **CONNECT request**: Sends the target hostname (type 0x03) and port.
   Receives the server's bound address and status.

## Limitations

- Only supports connect (not bind or UDP).
- Only supports domain name (type 0x03), not IPv4/IPv6.
- No support for GSSAPI authentication.
