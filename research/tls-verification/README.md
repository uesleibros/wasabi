# TLS Verification Research

This directory contains detailed notes on the TLS/SSL implementation in Wasabi,
focusing on certificate validation, Schannel configuration, and client certificates.

It is intended for maintainers who need to understand the low-level Windows APIs
(Crypt32, Secur32) and the rationale behind certain constants and calls.

## Related Documentation

For end-user documentation, see `../docs/README.md` (sections on `ValidateServerCert`,
`ClientCertThumb`, `ClientCertPfxPath`, etc.).
