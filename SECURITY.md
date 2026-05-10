# Security Policy

## Supported Versions

Only the latest release of Wasabi receives security updates. If you are on an older version, please update before reporting.

| Version | Supported |
|---|---|
| Latest release | ![](resources/svg/thumbsup.svg) |
| Older releases | ![](resources/svg/thumbsdown.svg) |

## Reporting a Vulnerability

If you discover a security vulnerability in Wasabi, **do not open a public GitHub issue**.

Use **[GitHub Private Vulnerability Reporting](https://github.com/uesleibros/wasabi/security/advisories/new)** for this repository instead.

When reporting, please include as much detail as you can:

- A clear description of the vulnerability
- Steps to reproduce it
- Affected version or commit
- Potential impact and attack surface
- Proof of concept, if you have one
- Suggested fix, if you have one

The more detail you provide, the faster the issue can be assessed and addressed.

## Scope

The following areas are in scope for security reports:

- TLS and Schannel handling (credential acquisition, handshake, encryption, decryption)
- WebSocket frame parsing (masking, length encoding, opcode handling, fragmentation)
- Buffer and memory safety (receive buffers, fragment reassembly, queue management)
- Input validation (URL parsing, header injection, proxy responses)
- Proxy handling and authentication (HTTP CONNECT, SOCKS5, NTLM)
- DNS resolution and connection logic
- Internal SHA-1 implementation used for the RFC 6455 handshake
- Random byte generation used for frame masking keys
- The async thunk and Win32 window subclassing mechanism
- Certificate validation logic (chain building, policy verification, revocation)
- Client certificate handling (PFX loading, Windows certificate store access)

The following are out of scope:

- Vulnerabilities in Windows itself or in the VBA/Office runtime
- Issues that require the attacker to already have code execution in the same process
- Social engineering attacks
- Misconfiguration of third-party servers or proxies
- Vulnerabilities in external servers that Wasabi connects to

If you are unsure whether something falls in scope, report it privately anyway and it will be assessed.

## Disclosure Policy

Wasabi follows a coordinated disclosure model.

Security issues are investigated privately. Once a fix is confirmed and released, the vulnerability may be disclosed publicly in the release notes or changelog, with credit to the reporter if desired.

There is no fixed timeline, but the goal is to move quickly. You will receive acknowledgment within a few days of the report and updates as the investigation progresses.

## Security Architecture Notes

Understanding what Wasabi does and does not do internally is useful context for security research.

**TLS** is handled entirely by Schannel through Windows SSPI (`secur32.dll`). Wasabi does not implement its own cipher suites, key exchange, or certificate validation algorithms. The security of TLS connections depends on the Schannel configuration of the host Windows installation, including available protocol versions and cipher suites.

**Certificate validation** is opt-in and disabled by default. When enabled via `WebSocketSetCertValidation`, Wasabi calls `CertGetCertificateChain` and `CertVerifyCertificateChainPolicy` with the SSL policy. Revocation checking via CRL and OCSP is a separate opt-in via `WebSocketSetRevocationCheck`.

**Frame masking keys** are generated using `BCryptGenRandom` from `bcrypt.dll` with the `BCRYPT_USE_SYSTEM_PREFERRED_RNG` flag. This replaced the previous use of `RtlGenRandom` (an undocumented `advapi32.dll` alias) as of v2.3.6.

**SHA-1** is used only for computing the `Sec-WebSocket-Accept` header during the RFC 6455 WebSocket handshake. It is not used for encryption, signing, or any security-sensitive purpose beyond what the WebSocket protocol itself requires.

**The async thunk** is a native machine-code stub allocated in executable memory (`VirtualAlloc` with `PAGE_EXECUTE_READ`). It subclasses a hidden Win32 window to receive `WSAAsyncSelect` socket events and dispatch them to VBA callbacks. This mechanism holds a pointer into the VBA runtime and must be explicitly cleaned up before the VBA project is reset.

**Proxy authentication** via NTLM uses the currently logged-on Windows user's credentials through SSPI (`secur32.dll`). No credentials are stored by Wasabi.
