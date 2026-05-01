# Test Certificates

This directory contains self-signed certificates and private keys used
exclusively for local TLS and mutual TLS (mTLS) testing during the
development of Wasabi.

> [!WARNING]
> These certificates are **not** issued by a public Certificate Authority.
> They must never be deployed to production systems, internal networks,
> or any environment that processes real data. All keys in this directory
> are intentionally short-lived and may be regenerated at any time.

## Contents

| File | Purpose |
|:---|:---|
| `ca.pem` | Self-signed root CA certificate (X.509 PEM) |
| `ca.key` | Private key of the test root CA |
| `server.pem` | Server certificate signed by the test CA |
| `server.key` | Private key of the server certificate |
| `client.pem` | Client certificate for mTLS handshake testing |
| `client.key` | Private key of the client certificate |
| `client.pfx` | PKCS#12 archive containing the client certificate, its private key, and the CA certificate |

> [!NOTE]
> The PKCS#12 file (`client.pfx`) is protected with the password `wasabi`.
> This password is hard-coded in the test suite and is not meant to be secret.

## Regeneration

Every certificate in this directory can be regenerated with OpenSSL.
Run the commands below from the `certs/` directory.

```powershell
# ----- Root CA -----------------------------------
openssl genrsa -out ca.key 2048
openssl req -x509 -new -nodes -key ca.key -sha256 -days 365 -out ca.pem -subj "/CN=Wasabi Test CA"

# ----- Server certificate -------------------------
openssl genrsa -out server.key 2048
openssl req -new -key server.key -out server.csr -subj "/CN=localhost"
openssl x509 -req -in server.csr -CA ca.pem -CAkey ca.key -CAcreateserial -out server.pem -days 365 -sha256

# ----- Client certificate -------------------------
openssl genrsa -out client.key 2048
openssl req -new -key client.key -out client.csr -subj "/CN=Wasabi Test Client"
openssl x509 -req -in client.csr -CA ca.pem -CAkey ca.key -CAcreateserial -out client.pem -days 365 -sha256

# ----- Client PFX (PKCS#12) -----------------------
openssl pkcs12 -export -out client.pfx -inkey client.key -in client.pem -certfile ca.pem -passout pass:wasabi

# ----- Clean up CSR files -------------------------
Remove-Item server.csr, client.csr
```

## Usage in Wasabi

### Server certificate validation

```vb
Dim h As Long
WebSocketSetCertValidation True, h
WebSocketConnect "wss://localhost:8443/ws", h
```

The connection will fail with `ERR_CERT_VALIDATE_FAILED` unless the test
CA has been explicitly trusted by the operating system (see below).

### Mutual TLS with client certificate

```vb
Dim h As Long
WebSocketSetClientCertPfx "resources/certs/client.pfx", "wasabi", h
WebSocketConnect "wss://localhost:8443/ws", h
```

> [!TIP]
> The path to the PFX file is relative to the VBA host document.
> Store the `certs/` folder alongside the workbook or provide an absolute
> path for reliable resolution.

## Trusting the Test CA on Windows

For server certificate validation to succeed during local tests, the test
CA must be installed in the Trusted Root Certification Authorities store.

1. Open `ca.pem` by double-clicking the file in Windows Explorer.
2. Select **Install Certificate**.
3. Choose **Local Machine** and proceed to the next step.
4. Select **Place all certificates in the following store**.
5. Browse to **Trusted Root Certification Authorities** and confirm.
6. Complete the wizard.

> [!IMPORTANT]
> Remove the test CA from the Trusted Root store as soon as testing is
> complete. Leaving a self-signed CA trusted on a development machine
> creates an unnecessary attack surface.
