# Server Certificate Validation

## Overview
Wasabi performs manual certificate validation only when `ValidateServerCert = True`.
The process follows these steps:

1. Retrieve the remote certificate from the security context (`QueryContextAttributes` with `SECPKG_ATTR_REMOTE_CERT_CONTEXT`).
2. Build a certificate chain using `CertGetCertificateChain`.
3. Set up an `SSL_EXTRA_CERT_CHAIN_POLICY_PARA` structure with the server hostname (`pwszServerName`).
4. Verify the chain with `CertVerifyCertificateChainPolicy` (policy `CERT_CHAIN_POLICY_SSL`).

## Why manual validation?
- Avoids classic WinInet/WinHTTP popups.
- Allows us to set the target hostname for proper SSL hostname matching.
- Gives us detailed error codes we can log or display.

## Revocation check
Optional, controlled by `EnableRevocationCheck`. When enabled, the flag `CERT_CHAIN_REVOCATION_CHECK_CHAIN` is passed to `CertGetCertificateChain`. This triggers a CRL or OCSP check. The connection will be rejected if the certificate is revoked or the revocation status cannot be determined (unless the offline flags are set on the credential – see `schannel-constants.md`).

## Known issues
- The `dwError` field in `CERT_CHAIN_POLICY_STATUS` is not a Win32 error code; it's a chain policy error (e.g., `TRUST_E_CERT_SIGNATURE`). The code currently reports it as hex.
- If the server uses a self-signed certificate, validation will fail. The user must either disable validation or add the CA certificate to the trusted store (not handled by Wasabi).

## References
- [Microsoft: Certificate Chain Validation](https://learn.microsoft.com/en-us/windows/win32/seccrypto/certificate-chain-validation)
- `CertGetCertificateChain`, `CertVerifyCertificateChainPolicy`, `SSL_EXTRA_CERT_CHAIN_POLICY_PARA`
