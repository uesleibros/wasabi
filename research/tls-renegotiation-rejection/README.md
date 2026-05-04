# TLS Renegotiation Rejection

Wasabi deliberately terminates any connection where the server requests a TLS
renegotiation. This document explains the rationale and the implementation.

## Reference

- `TLSDecrypt()` and the constant `SEC_I_RENEGOTIATE` in `Wasabi.bas`.
