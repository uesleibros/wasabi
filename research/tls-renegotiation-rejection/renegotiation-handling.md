# TLS Renegotiation Handling

## Detection

Inside `TLSDecrypt()`, after calling `DecryptMessage`, the return code is
checked. If it equals `SEC_I_RENEGOTIATE (&H90321)`, the server has requested a
TLS renegotiation.

```vb
If result = SEC_I_RENEGOTIATE Then
    SetError ERR_TLS_RENEGOTIATE, "TLS renegotiation requested - closing", _
              "Secure connection interrupted (renegotiation).", handle, SEC_I_RENEGOTIATE
    .Connected = False
    If .AutoReconnect Then TryReconnect handle
    Exit Sub
End If
```

## Why Wasabi rejects renegotiation

1. **Complexity and security risk:** TLS renegotiation has a history of
   vulnerabilities (e.g., CVE-2009-3555). While secure renegotiation (RFC 5746)
   addresses those issues, implementing it correctly in a VBA module that already
   manages handshake, encryption, and certificate validation would add
   significant complexity.

2. **No application‑level need:** In the context of WebSocket, servers rarely
   require renegotiation after the initial handshake. If a server does, it is
   almost certainly a misconfiguration or an attempted attack.

3. **Simplicity:** Terminating the connection immediately and allowing the user’s
   auto‑reconnect logic (if enabled) to establish a fresh, clean session is a
   robust and easy‑to‑understand behaviour.

## Error propagation

- `ERR_TLS_RENEGOTIATE` is set in the connection’s `LastError` field.
- The connection’s `Connected` flag is set to `False`.
- If `AutoReconnect` is enabled, `TryReconnect` is called immediately to
  re‑establish the session.

## References

- Microsoft documentation: [`SEC_I_RENEGOTIATE`](https://learn.microsoft.com/en-us/windows/win32/secauthn/sspi-status-codes)
- RFC 5746: Transport Layer Security (TLS) Renegotiation Indication Extension
