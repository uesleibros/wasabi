# NTLM Authentication for Proxies

`GenerateNtlmToken(handle, proxyAuthHeader, proxyHost)` generates NTLM
authentication tokens using the SSPI (`secur32.dll`).

## Motivation

Many corporate proxies require NTLM (Windows Integrated Authentication).
VBA has no built‑in support for NTLM, so Wasabi implements it via the
SSPI API, similar to how it does TLS with Schannel.

## Flow

1. **Acquire credentials**: `AcquireCredentialsHandle` with the "NTLM"
   package, outbound.

2. **Generate Type 1**: A hardcoded Type 1 token (base64) is sent
   initially to avoid a round trip. This token is standard and
   indicates NTLMv2 capability.

3. **Server challenge**: The proxy returns a 407 response with a
   `Proxy-Authenticate: NTLM <challenge>` header.

4. **Generate Type 3**: The challenge is extracted, base64‑decoded,
   and passed to `InitializeSecurityContextContinue` which generates
   the Type 3 (authenticate) token.

5. **Resend CONNECT**: The Type 3 token is sent in a new
   `Proxy-Authorization` header, completing the handshake.

## Handles

A dedicated `hNtlmCred` handle is stored in the connection structure,
separate from the TLS credential handle. It is freed in
`FreeSecurityHandles`.

## Limitations

- Only supports NTLMv2 (not Kerberos).
- The hardcoded Type 1 token may not work with all proxies (though it
  is widely compatible).
- No session security (encryption) after authentication – this is just
  for tunnel establishment.
