# Schannel Constants Reference

## Protocol versions
| Constant | Value | Notes |
|---|---|---|
| `SP_PROT_TLS1_2_CLIENT` | `&H800` | Required for most modern servers. |
| `SP_PROT_TLS1_3_CLIENT` | `&H2000` | Attempted if available, but schannel silently falls back. |

## Credential flags (`schannelCred.dwFlags`)
| Flag | Value | Why |
|---|---|---|
| `SCH_CRED_NO_DEFAULT_CREDS` | `&H10` | Prevent automatic use of the client's (often personal) credentials. |
| `SCH_CRED_MANUAL_CRED_VALIDATION` | `&H8` | We handle certificate validation ourselves via `CertGetCertificateChain`/`CertVerifyCertificateChainPolicy`. |
| `SCH_CRED_IGNORE_NO_REVOCATION_CHECK` | `&H800` | Allow revocation check to be skipped if the CDP is offline (avoids hard failures on misconfigured servers). |
| `SCH_CRED_IGNORE_REVOCATION_OFFLINE` | `&H1000` | Same as above, more specific to offline revocation. |

Note: The combination `SCH_CRED_NO_DEFAULT_CREDS | SCH_CRED_MANUAL_CRED_VALIDATION` gives us full control, which is necessary because the default Windows chain validation can trigger popups or unwanted behavior in VBA.

## Missing constant bug (v2.1.1-vNext)
`CERT_CHAIN_REVOCATION_CHECK_CHAIN` (`&H20000000`) was not declared, causing a compile error when `EnableRevocationCheck = True`. Added alongside the other crypt32 constants.

## Context flags (`ISC_REQ_*`)
- `ISC_REQ_SEQUENCE_DETECT`, `ISC_REQ_REPLAY_DETECT`, `ISC_REQ_CONFIDENTIALITY`, `ISC_REQ_STREAM` – standard for TLS.
- `ISC_REQ_ALLOCATE_MEMORY` – required so that the output token is allocated by schannel (we must free it with `FreeContextBuffer`).

## Why not use `SCH_CRED_REVOCATION_CHECK_CHAIN`?
Because it's applied at the schannel level, which can cause connection termination without our control. Instead, we use manual validation and apply `CERT_CHAIN_REVOCATION_CHECK_CHAIN` only when the user requests it, and we handle the error gracefully.
