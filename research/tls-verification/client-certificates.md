# Client Certificate Support

Wasabi supports two methods:

1. **Thumbprint (Subject string)** – `ClientCertThumb`. Searches the current user's "MY" store.
2. **PFX file (PKCS#12)** – `ClientCertPfxPath` + `ClientCertPfxPass`. Imported into a temporary in-memory store.

## Implementation details
- `CertOpenStore(CERT_STORE_PROV_SYSTEM, ... MY store)` for thumbprint.
- `PFXImportCertStore` for PFX. Flags: `CRYPT_EXPORTABLE` (required for Schannel) and `PKCS12_ALLOW_OVERWRITE_KEY`.
- The resulting `pClientCertCtx` is stored in `m_ClientCertContextPtrs(handle)` (as `LongPtr` on VBA7, `Long` on VBA6). This array is necessary because `SCHANNEL_CRED.paCred` expects an array of pointers.
- In `ConnectHandle`, after loading the certificate, we set `schannelCred.cCreds = 1` and `schannelCred.paCred = VarPtr(m_ClientCertContextPtrs(handle))`. This is why the array is required – we need a stable memory address pointing to the context.

## Potential pitfalls
- The `CERT_FIND_SUBJECT_STR_A` flag matches the entire subject string, which may be inconsistent. In some environments, matching by thumbprint (hash) is more reliable, but that requires `CERT_FIND_SHA1_HASH` or similar. We currently rely on the subject string containing the thumbprint (as displayed in Windows certificate manager) – this is ambiguous.
- PFX password: passed as a `LongPtr` to a null-terminated Unicode string (or `NULL_PTR` if empty). Mistyped password will cause `PFXImportCertStore` to fail with `NTE_FAIL`.
- Certificate store cleanup: the temporary store (`hClientCertStore`) must be closed with `CertCloseStore`, and the certificate context freed with `CertFreeCertificateContext`. Done in `FreeSecurityHandles`.

## Future improvements
- Add support for certificate thumbprint as hex string (more precise).
- Support for PEM certificates.
