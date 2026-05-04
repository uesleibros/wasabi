# NTLM Proxy Authentication

## Motivation
HTTP proxies requiring NTLM authentication cannot use simple Basic auth. Wasabi implements a custom NTLM handshake using the SSPI (`secur32.dll`), similar to TLS but with the "NTLM" package.

## Process
1. Client sends a CONNECT request with `Proxy-Authorization: NTLM <Type1 token>` (hardcoded as a known Type1).
2. Proxy responds with `407 Proxy Authentication Required` and a `Proxy-Authenticate: NTLM <challenge>` header.
3. Wasabi extracts the challenge, passes it to `InitializeSecurityContextContinue` to generate a Type3 token.
4. Client resends the CONNECT with the Type3 token.

## Constants and handles
- `SECPKG_CRED_OUTBOUND_NTLM = &H2` (misleading, same as `SECPKG_CRED_OUTBOUND` but we use a separate variable for clarity).
- A dedicated `hNtlmCred` handle is stored in the connection structure, and freed in `FreeSecurityHandles`.

## Known limitations
- Only works for proxy authentication; not for direct connections.
- The hardcoded Type1 token may not work with all proxies (though it's standard).
- No support for NTLMv2 session security; this is just for tunnel establishment.

## References
- `GenerateNtlmToken` in `Wasabi.bas`
